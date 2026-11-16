/**
 * ping-online.js
 * ระบบจัดการห้องแข่งขันออนไลน์ 1 ต่อ 1 ด้วย Firebase Realtime Database + PeerJS WebRTC DataChannel
 * รองรับ Host-Authoritative Gameplay, 3D Spatial Audio Mirroring, Latency Compensation และ Screen Reader Accessibility
 */
import {
  ref,
  get,
  set,
  update,
  remove,
  onValue,
  onDisconnect,
  runTransaction
} from "https://www.gstatic.com/firebasejs/10.8.0/firebase-database.js";
import { database, GAME_ID, encodeUserKey } from "./auth.js";

/**
 * ตั้งค่า ICE / STUN / TURN Server สำหรับ PeerJS WebRTC
 * หมายเหตุสำคัญ:
 * - STUN มาตรฐานใช้สำหรับเครือข่ายทั่วไปและ Wi-Fi
 * - สำหรับเครือข่ายมือถือที่มี Symmetric NAT เข้มงวด (เช่น 4G/5G บางค่าย) ผู้พัฒนาต้องใส่ TURN Server จริงที่ใช้งานได้ใน TURN_CONFIG ด้านล่าง
 * - ห้ามใส่ Credential ปลอม หากยังไม่มี TURN Server ให้เว้นว่างไว้ ระบบจะตรวจสอบ Timeout และแจ้งเตือนผู้เล่นอย่างชัดเจนเมื่อเชื่อมต่อข้ามเครือข่ายไม่สำเร็จ
 */
export const WEBRTC_ICE_SERVERS = [
  { urls: "stun:stun.l.google.com:19302" },
  { urls: "stun:stun1.l.google.com:19302" },
  { urls: "stun:stun2.l.google.com:19302" }
  // หากมี TURN Server จริง ให้เพิ่มอ็อบเจกต์รูปแบบด้านล่างนี้:
  // {
  //   urls: "turn:your-turn-server.example.com:3478",
  //   username: "YOUR_TURN_USERNAME",
  //   credential: "YOUR_TURN_PASSWORD"
  // }
];

const ONLINE_CONFIG = {
  SERVE_TIMEOUT_MS: 6000,
  BASE_HIT_WINDOW_MS: 1150,
  MAX_LATENCY_COMPENSATION_MS: 450,
  HIGH_PING_PAUSE_THRESHOLD_MS: 650,
  POINT_TRANSITION_MS: 2400,
  RECONNECT_GRACE_MS: 30000,
  PEER_CONNECT_TIMEOUT_MS: 15000,
  PING_INTERVAL_MS: 2000
};

/**
 * แปลงทิศทางมุมมองของผู้เล่นฝั่งตรงข้าม (Mirror Perspective)
 * เมื่อผู้เล่น A ตีไปทางซ้ายของตัวเอง ('left') ลูกจะวิ่งเข้าหาฝั่งขวาของผู้เล่น B ('right')
 */
export function mirrorDirection(dir) {
  if (dir === 'left') return 'right';
  if (dir === 'right') return 'left';
  return 'center';
}

export class OnlineMatchController {
  constructor(options) {
    this.sound = options.sound;
    this.motion = options.motion;
    this.directions = options.directions;
    this.flightDurations = options.flightDurations;
    this.getCurrentUser = options.getCurrentUser;
    this.addAccessibleClickListener = options.addAccessibleClickListener;
    this.onExitToHome = options.onExitToHome;
    this.ui = options.ui;

    this.db = database;
    this.activeMode = false;
    this.inMatch = false;

    this.currentRoomId = null;
    this.role = null; // 'host' | 'client'
    this.matchId = null;

    this.peer = null;
    this.conn = null;
    this.peerReady = false;
    this.channelReady = false;
    this.peerConnectTimer = null;

    this.roomUnsubscribe = null;
    this.lobbyUnsubscribe = null;
    this.presenceDisconnectRef = null;

    this.rttMs = 80;
    this.pingTimer = null;
    this.reconnectTimer = null;
    this.reconnectCountdownInterval = null;
    this.opponentDisconnected = false;
    this.lastJoinedClientKey = null;

    // สถานะการแข่งขัน (Host-Authoritative)
    this.matchStatus = 'idle'; // 'idle' | 'playing' | 'between_sets' | 'paused_disconnect' | 'finished' | 'aborted'
    this.hostUsername = '';
    this.clientUsername = '';
    this.hostSetsWon = 0;
    this.clientSetsWon = 0;
    this.currentSetNumber = 1;
    this.hostScore = 0;
    this.clientScore = 0;
    this.initialServer = 'host';
    this.currentServer = 'host';

    // สถานะ Rally ปัจจุบัน
    this.rallyId = 0;
    this.shotSeq = 0;
    this.rallyCount = 0;
    this.canLocalPlayerHit = false;
    this.expectedLocalDir = 'center';
    this.localBouncePerfTime = 0;
    this.activeRallyTimer = null;
    this.hostPendingReceive = null;

    // คิวประกาศสำหรับ Screen Reader
    this.announceQueue = [];
    this.isAnnouncing = false;
    this.announceClearTimer = null;

    this.bindLobbyEvents();
  }

  bindLobbyEvents() {
    const {
      createRoomBtn,
      refreshRoomsBtn,
      backFromLobbyBtn,
      startOnlineMatchBtn,
      leaveRoomBtn
    } = this.ui;

    this.addAccessibleClickListener(createRoomBtn, async () => {
      await this.requestSensorAndAudioOnGesture();
      await this.createRoom();
    });

    this.addAccessibleClickListener(refreshRoomsBtn, () => {
      this.announcePolite("กำลังอัปเดตรายการห้องออนไลน์");
      this.RefreshRoomsManual();
    });

    this.addAccessibleClickListener(backFromLobbyBtn, async () => {
      await this.closeLobbyAndExit();
    });

    this.addAccessibleClickListener(startOnlineMatchBtn, async () => {
      await this.requestSensorAndAudioOnGesture();
      await this.hostStartMatch();
    });

    this.addAccessibleClickListener(leaveRoomBtn, async () => {
      await this.leaveCurrentRoom(true);
    });
  }

  /**
   * ขออนุญาต Motion Sensor และปลุก AudioContext ทันทีภายใต้ User Gesture
   */
  async requestSensorAndAudioOnGesture() {
    const permPromise = this.motion.requestPermission();
    const soundPromise = this.sound.init();
    const [granted] = await Promise.all([permPromise, soundPromise]);
    this.sound.playBGM(0.20);
    if (!granted && this.ui.permissionMsg) {
      this.ui.permissionMsg.textContent = "ระบบไม่สามารถเข้าถึง Motion Sensor ได้ คุณสามารถใช้คีย์บอร์ดเพื่อควบคุมแทนได้";
      this.ui.permissionMsg.classList.remove('hidden');
    } else if (this.ui.permissionMsg) {
      this.ui.permissionMsg.classList.add('hidden');
    }
    return granted;
  }

  /**
   * ระบบจัดการคิวประกาศเสียงสำหรับ Screen Reader เพื่อไม่ให้ตัดคำสำคัญและไม่รบกวนเสียงลูกปิงปองระหว่างตี
   * ปรับปรุงให้ประกาศผ่านกล่องสถานะหลักเพียงจุดเดียว ป้องกัน VoiceOver อ่านซ้ำ 2 ครั้ง
   */
  announcePolite(message, forceDuringRally = false) {
    if (!message) return;
    if (this.inMatch && this.matchStatus === 'playing' && this.rallyCount > 0 && !forceDuringRally) {
      return;
    }
    this.announceQueue.push(message);
    this.processAnnounceQueue();
  }

  processAnnounceQueue() {
    if (this.isAnnouncing || this.announceQueue.length === 0) return;
    this.isAnnouncing = true;
    const nextMsg = this.announceQueue.shift();

    if (this.ui.onlineSrAnnouncer) {
      this.ui.onlineSrAnnouncer.textContent = '';
    }

    if (this.inMatch && this.ui.gameStatus) {
      if (this.ui.gameStatus.textContent === nextMsg) {
        this.ui.gameStatus.textContent = '';
        setTimeout(() => {
          if (this.ui.gameStatus) this.ui.gameStatus.textContent = nextMsg;
        }, 20);
      } else {
        this.ui.gameStatus.textContent = nextMsg;
      }
    } else if (this.ui.lobbyStatusBox) {
      if (this.ui.lobbyStatusBox.textContent === nextMsg) {
        this.ui.lobbyStatusBox.textContent = '';
        setTimeout(() => {
          if (this.ui.lobbyStatusBox) this.ui.lobbyStatusBox.textContent = nextMsg;
        }, 20);
      } else {
        this.ui.lobbyStatusBox.textContent = nextMsg;
      }
    } else if (this.ui.onlineSrAnnouncer) {
      this.ui.onlineSrAnnouncer.textContent = nextMsg;
    }

    clearTimeout(this.announceClearTimer);
    this.announceClearTimer = setTimeout(() => {
      this.isAnnouncing = false;
      if (this.announceQueue.length > 0) {
        this.processAnnounceQueue();
      }
    }, 450);
  }

  async openLobby() {
    this.activeMode = true;
    this.inMatch = false;
    this.ui.homeScreen.classList.add('hidden');
    this.ui.rulesScreen.classList.add('hidden');
    this.ui.roundSelectScreen.classList.add('hidden');
    this.ui.leaderboardScreen.classList.add('hidden');
    this.ui.gameScreen.classList.add('hidden');

    this.ui.startBtn.classList.add('hidden');
    this.ui.onlineModeBtn.classList.add('hidden');
    this.ui.selectRoundBtn.classList.add('hidden');
    this.ui.rulesBtn.classList.add('hidden');
    this.ui.inGameLeaderboardBtn.classList.add('hidden');

    this.ui.onlineLobbyScreen.classList.remove('hidden');
    this.ui.backFromLobbyBtn.classList.remove('hidden');

    await this.loadPlayerOnlineStatsUI();
    this.subscribeWaitingRooms();

    // ตรวจสอบว่าผู้เล่นมีห้องที่กำลังเปิดค้างไว้หรือไม่ (เช่น กรณีกด Refresh)
    await this.tryRecoverExistingRoom();

    setTimeout(() => {
      try {
        this.ui.onlineLobbyScreen.querySelector('h1').focus();
      } catch (e) {}
    }, 100);
  }

  async loadPlayerOnlineStatsUI() {
    const user = this.getCurrentUser();
    if (!user || !this.ui.onlineStatsBadge) return;
    const key = encodeUserKey(user.username);
    try {
      const snap = await get(ref(this.db, `${GAME_ID}/onlineStats/${key}`));
      if (snap.exists()) {
        const st = snap.val();
        this.ui.onlineStatsBadge.textContent = `สถิติออนไลน์: แข่ง ${st.totalMatches || 0} | ชนะ ${st.wins || 0} | แพ้ ${st.losses || 0} | ยุติ ${st.aborted || 0}`;
      } else {
        this.ui.onlineStatsBadge.textContent = `สถิติออนไลน์: แข่ง 0 | ชนะ 0 | แพ้ 0 | ยุติ 0`;
      }
    } catch (e) {
      this.ui.onlineStatsBadge.textContent = `สถิติออนไลน์: พร้อมแข่งขัน`;
    }
  }

  async closeLobbyAndExit() {
    await this.leaveCurrentRoom(true);
    this.unsubscribeWaitingRooms();
    this.activeMode = false;
    this.inMatch = false;

    this.ui.onlineLobbyScreen.classList.add('hidden');
    this.ui.backFromLobbyBtn.classList.add('hidden');
    this.ui.leaveRoomBtn.classList.add('hidden');
    this.ui.startOnlineMatchBtn.classList.add('hidden');

    if (typeof this.onExitToHome === 'function') {
      this.onExitToHome();
    }
  }

  subscribeWaitingRooms() {
    this.unsubscribeWaitingRooms();
    const roomsRef = ref(this.db, `${GAME_ID}/onlineRooms`);
    this.lobbyUnsubscribe = onValue(roomsRef, (snapshot) => {
      this.renderWaitingRoomsList(snapshot.exists() ? snapshot.val() : {});
    });
  }

  unsubscribeWaitingRooms() {
    if (this.lobbyUnsubscribe) {
      this.lobbyUnsubscribe();
      this.lobbyUnsubscribe = null;
    }
  }

  RefreshRoomsManual() {
    get(ref(this.db, `${GAME_ID}/onlineRooms`)).then((snap) => {
      this.renderWaitingRoomsList(snap.exists() ? snap.val() : {});
    }).catch(() => {});
  }

  renderWaitingRoomsList(roomsObj) {
    const container = this.ui.waitingRoomsContainer;
    if (!container) return;
    container.innerHTML = '';

    const entries = Object.values(roomsObj || {}).filter((r) => {
      return r && r.roomId && r.status === 'waiting' && r.host && !r.client;
    });

    if (entries.length === 0) {
      const emptyEl = document.createElement('p');
      emptyEl.className = 'desc';
      emptyEl.setAttribute('tabindex', '0');
      emptyEl.textContent = 'ยังไม่มีห้องที่กำลังรอผู้เล่น กดปุ่มสร้างห้องใหม่เพื่อเริ่มได้เลย';
      container.appendChild(emptyEl);
      return;
    }

    entries.sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));

    entries.forEach((room) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'btn-secondary';
      btn.textContent = `ห้อง ${room.roomId} - ผู้สร้าง: ${room.host.username} (รอผู้เล่น 1/2)`;
      btn.setAttribute('aria-label', `เข้าร่วมห้อง ${room.roomId} สร้างโดย ${room.host.username}`);

      if (this.currentRoomId) {
        btn.disabled = true;
      }

      this.addAccessibleClickListener(btn, async () => {
        await this.requestSensorAndAudioOnGesture();
        await this.joinRoom(room.roomId);
      });

      container.appendChild(btn);
    });
  }

  /**
   * ตรวจสอบและกู้คืนห้องกรณีผู้เล่นกด Refresh หน้าเว็บภายในเวลา 30 วินาที
   */
  async tryRecoverExistingRoom() {
    const user = this.getCurrentUser();
    if (!user) return;
    const savedRoomId = sessionStorage.getItem('pingpong_active_online_room');
    if (!savedRoomId) return;

    try {
      const roomRef = ref(this.db, `${GAME_ID}/onlineRooms/${savedRoomId}`);
      const snap = await get(roomRef);
      if (!snap.exists()) {
        sessionStorage.removeItem('pingpong_active_online_room');
        return;
      }
      const room = snap.val();
      if (room.status === 'finished' || room.status === 'aborted') {
        sessionStorage.removeItem('pingpong_active_online_room');
        return;
      }

      const myKey = encodeUserKey(user.username);
      if (room.host && room.host.userKey === myKey) {
        this.currentRoomId = savedRoomId;
        this.role = 'host';
        this.matchId = room.matchId;
        this.lastJoinedClientKey = room.client ? room.client.userKey : null;
        await this.initPeerAndBindRoom(savedRoomId, 'host');
        this.announcePolite(`กู้คืนห้อง ${savedRoomId} สำเร็จในฐานะหัวหน้าห้อง`, true);
      } else if (room.client && room.client.userKey === myKey) {
        this.currentRoomId = savedRoomId;
        this.role = 'client';
        this.matchId = room.matchId;
        this.lastJoinedClientKey = room.client.userKey;
        await this.initPeerAndBindRoom(savedRoomId, 'client');
        this.announcePolite(`กู้คืนห้อง ${savedRoomId} สำเร็จในฐานะผู้ท้าชิง`, true);
      } else {
        sessionStorage.removeItem('pingpong_active_online_room');
      }
    } catch (e) {
      console.warn("ไม่สามารถกู้คืนห้องเดิมได้:", e);
    }
  }

  /**
   * สร้างรหัสห้องรูปแบบ PingPong00001 เพิ่มตามลำดับแบบ Atomic ด้วย Firebase Transaction
   */
  async generateNextRoomId() {
    const counterRef = ref(this.db, `${GAME_ID}/onlineRoomCounter`);
    const result = await runTransaction(counterRef, (currentValue) => {
      return (currentValue || 0) + 1;
    });
    if (!result.committed) {
      throw new Error("ไม่สามารถจองหมายเลขห้องได้ กรุณาลองอีกครั้ง");
    }
    const seqNum = result.snapshot.val();
    return `PingPong${String(seqNum).padStart(5, '0')}`;
  }

  async createRoom() {
    const user = this.getCurrentUser();
    if (!user) return;
    if (this.currentRoomId) {
      this.announcePolite("คุณอยู่ในห้องเรียบร้อยแล้ว");
      return;
    }

    try {
      this.announcePolite("กำลังสร้างห้องแข่งขันออนไลน์...");
      const roomId = await this.generateNextRoomId();
      const myKey = encodeUserKey(user.username);
      const matchId = `${roomId}_${Date.now()}`;

      const roomData = {
        roomId: roomId,
        matchId: matchId,
        status: 'waiting',
        createdAt: Date.now(),
        updatedAt: Date.now(),
        playerCount: 1,
        host: {
          username: user.username,
          userKey: myKey,
          peerId: '',
          connected: true,
          joinedAt: Date.now()
        },
        client: null,
        matchState: {
          hostSetsWon: 0,
          clientSetsWon: 0,
          currentSetNumber: 1,
          hostScore: 0,
          clientScore: 0,
          initialServer: 'host',
          currentServer: 'host'
        }
      };

      const roomRef = ref(this.db, `${GAME_ID}/onlineRooms/${roomId}`);
      await set(roomRef, roomData);

      this.currentRoomId = roomId;
      this.role = 'host';
      this.matchId = matchId;
      this.lastJoinedClientKey = null;
      sessionStorage.setItem('pingpong_active_online_room', roomId);

      // 1. เล่นเสียง create.mp3 เมื่อกดสร้างห้องสำเร็จ (เฉพาะฝั่ง Host)
      this.sound.play('create', 0, 1.0, false);

      await this.initPeerAndBindRoom(roomId, 'host');
      this.announcePolite(`สร้างห้องสำเร็จ รหัสห้อง ${roomId} กำลังรอผู้เล่นคนที่สองเข้าร่วม`);
    } catch (err) {
      console.error(err);
      this.announcePolite(err.message || "เกิดข้อผิดพลาดในการสร้างห้อง");
    }
  }

  /**
   * เข้าร่วมห้องโดยใช้ Firebase Transaction ป้องกันผู้เล่นคนที่ 3 หรือการกดพร้อมกัน (Race Condition)
   */
  async joinRoom(roomId) {
    const user = this.getCurrentUser();
    if (!user) return;
    if (this.currentRoomId) {
      this.announcePolite("คุณอยู่ในห้องอื่นแล้ว กรุณาออกจากห้องเดิมก่อน");
      return;
    }

    const myKey = encodeUserKey(user.username);
    const roomRef = ref(this.db, `${GAME_ID}/onlineRooms/${roomId}`);

    try {
      this.announcePolite(`กำลังเข้าร่วมห้อง ${roomId}...`);
      let rejectReason = "";

      const txResult = await runTransaction(roomRef, (room) => {
        if (room === null) {
          return room;
        }
        if (room.host && room.host.userKey === myKey) {
          rejectReason = "คุณเป็นเจ้าของห้องนี้อยู่แล้ว";
          return;
        }
        if (room.status !== 'waiting' || (room.client && room.client.userKey && room.client.userKey !== myKey)) {
          rejectReason = "ห้องนี้มีผู้เล่นครบ 2 คนแล้ว หรือปิดรับผู้เล่นแล้ว";
          return;
        }

        room.client = {
          username: user.username,
          userKey: myKey,
          peerId: '',
          connected: true,
          joinedAt: Date.now()
        };
        room.playerCount = 2;
        room.status = 'ready';
        room.updatedAt = Date.now();
        return room;
      });

      if (!txResult.committed || !txResult.snapshot.exists()) {
        this.announcePolite(rejectReason || "ไม่สามารถเข้าร่วมห้องนี้ได้ ห้องอาจเต็มหรือถูกปิดไปแล้ว");
        return;
      }

      const updatedRoom = txResult.snapshot.val();
      if (!updatedRoom.client || updatedRoom.client.userKey !== myKey) {
        this.announcePolite("ห้องเต็มแล้ว ไม่สามารถเข้าร่วมเป็นคนที่สามได้");
        return;
      }

      this.currentRoomId = roomId;
      this.role = 'client';
      this.matchId = updatedRoom.matchId;
      this.lastJoinedClientKey = myKey;
      sessionStorage.setItem('pingpong_active_online_room', roomId);

      // 2. เล่นเสียง select.mp3 เมื่อกดเข้าร่วมห้องสำเร็จ (ฝั่ง Client)
      this.sound.play('select', 0, 1.0, false);

      await this.initPeerAndBindRoom(roomId, 'client');
      this.announcePolite(`เข้าร่วมห้อง ${roomId} สำเร็จ พบกับ ${updatedRoom.host.username} กำลังเชื่อมต่อสัญญาณตรง`);
    } catch (err) {
      console.error(err);
      this.announcePolite("ไม่สามารถเข้าร่วมห้องได้ กรุณาลองเลือกห้องอื่น");
    }
  }

  /**
   * สร้าง PeerJS Instance และผูก Listener กับห้องใน Firebase
   */
  async initPeerAndBindRoom(roomId, role) {
    this.cleanupPeerOnly();
    this.updateRoomPanelVisibility(true);

    const PeerClass = window.Peer;
    if (!PeerClass) {
      this.announcePolite("ไม่พบไลบรารี PeerJS กรุณาตรวจสอบการเชื่อมต่ออินเทอร์เน็ต");
      return;
    }

    this.peer = new PeerClass(undefined, {
      config: {
        iceServers: WEBRTC_ICE_SERVERS
      }
    });

    this.peer.on('open', async (peerId) => {
      this.peerReady = true;
      const peerPath = `${GAME_ID}/onlineRooms/${roomId}/${role}`;
      await update(ref(this.db, peerPath), {
        peerId: peerId,
        connected: true,
        lastSeen: Date.now()
      });

      const connFieldRef = ref(this.db, `${peerPath}/connected`);
      this.presenceDisconnectRef = onDisconnect(connFieldRef);
      await this.presenceDisconnectRef.set(false);
    });

    this.peer.on('connection', (incomingConn) => {
      if (this.role === 'host') {
        this.setupDataConnection(incomingConn);
      }
    });

    this.peer.on('error', (err) => {
      console.warn("PeerJS error:", err);
      this.announcePolite("สัญญาณ PeerJS ขัดข้อง กำลังพยายามตรวจสอบการเชื่อมต่อ");
    });

    this.subscribeRoomChanges(roomId);
  }

  subscribeRoomChanges(roomId) {
    if (this.roomUnsubscribe) {
      this.roomUnsubscribe();
      this.roomUnsubscribe = null;
    }

    const roomRef = ref(this.db, `${GAME_ID}/onlineRooms/${roomId}`);
    this.roomUnsubscribe = onValue(roomRef, (snap) => {
      if (!snap.exists()) {
        if (this.currentRoomId === roomId && !this.inMatch) {
          this.announcePolite("ห้องถูกปิดแล้ว");
          this.leaveCurrentRoom(false);
        }
        return;
      }
      const room = snap.val();
      this.handleRoomSnapshot(room);
    });
  }

  handleRoomSnapshot(room) {
    if (!room) return;
    this.hostUsername = room.host ? room.host.username : '-';
    this.clientUsername = room.client ? room.client.username : 'รอผู้เล่นเข้าร่วม...';

    // 2. เล่นเสียง select.mp3 บนฝั่ง Host ทันทีเมื่อมีผู้เล่นกดเข้าร่วมห้องสำเร็จ
    if (this.role === 'host') {
      const currentClientKey = (room.client && room.client.userKey) ? room.client.userKey : null;
      if (currentClientKey && currentClientKey !== this.lastJoinedClientKey) {
        this.lastJoinedClientKey = currentClientKey;
        this.sound.play('select', 0, 1.0, false);
      } else if (!currentClientKey) {
        this.lastJoinedClientKey = null;
      }
    }

    // อัปเดตข้อมูลในหน้า Lobby
    if (this.ui.roomCodeDisplay) {
      this.ui.roomCodeDisplay.textContent = `รหัสห้อง: ${room.roomId}`;
    }
    if (this.ui.roomHostDisplay) {
      this.ui.roomHostDisplay.textContent = `หัวหน้าห้อง (Host): ${this.hostUsername}`;
    }
    if (this.ui.roomClientDisplay) {
      this.ui.roomClientDisplay.textContent = `ผู้ท้าชิง (Client): ${this.clientUsername}`;
    }

    const hasBothPlayers = Boolean(room.host && room.client && room.playerCount === 2);

    // หากผู้เล่นเป็น Client และเห็น Peer ID ของ Host แล้ว แต่ยังไม่ได้เชื่อม DataChannel ให้เริ่มเชื่อมทันที
    if (
      this.role === 'client' &&
      this.peerReady &&
      !this.conn &&
      room.host &&
      room.host.peerId &&
      room.host.connected
    ) {
      this.connectToHostPeer(room.host.peerId);
    }

    // ตรวจสอบการหลุดการเชื่อมต่อของคู่แข่งจาก Firebase Presence
    const oppNode = this.role === 'host' ? room.client : room.host;
    if (this.inMatch && this.matchStatus !== 'finished' && this.matchStatus !== 'aborted') {
      if (oppNode && oppNode.connected === false && !this.opponentDisconnected) {
        this.handleOpponentDisconnect();
      } else if (oppNode && oppNode.connected === true && this.opponentDisconnected) {
        if (this.role === 'client' && room.host && room.host.peerId && (!this.conn || !this.channelReady)) {
          this.connectToHostPeer(room.host.peerId);
        }
      }
    }

    // อัปเดตสถานะในหน้าห้องรอ
    if (!this.inMatch) {
      if (!hasBothPlayers) {
        this.ui.roomConnectionDisplay.textContent = 'สถานะห้อง: รอผู้เล่นคนที่ 2 เข้าร่วม (1/2)';
        this.ui.startOnlineMatchBtn.classList.add('hidden');
      } else if (!this.channelReady) {
        this.ui.roomConnectionDisplay.textContent = 'สถานะห้อง: ผู้เล่นครบ 2 คนแล้ว กำลังเชื่อมสัญญาณ PeerJS...';
        this.ui.startOnlineMatchBtn.classList.add('hidden');
      } else {
        this.ui.roomConnectionDisplay.textContent = `สถานะห้อง: เชื่อมต่อพร้อมแข่งขัน (Ping ${Math.round(this.rttMs)} ms)`;
        if (this.role === 'host') {
          this.ui.startOnlineMatchBtn.classList.remove('hidden');
        } else {
          this.ui.startOnlineMatchBtn.classList.add('hidden');
        }
      }
    }

    // กรณีห้องถูกบันทึกว่ายุติการแข่งขันจากอีกฝ่าย
    if (room.status === 'aborted' && this.matchStatus !== 'aborted' && this.matchStatus !== 'finished') {
      this.finishMatchAsAborted(room.abortReason || "การแข่งขันยุติลงเนื่องจากปัญหาการเชื่อมต่อ");
    }
  }

  connectToHostPeer(hostPeerId) {
    if (!this.peer || !hostPeerId) return;
    clearTimeout(this.peerConnectTimer);

    const outgoingConn = this.peer.connect(hostPeerId, {
      reliable: true,
      serialization: 'json'
    });

    this.setupDataConnection(outgoingConn);

    this.peerConnectTimer = setTimeout(() => {
      if (!this.channelReady) {
        this.announcePolite(
          "ไม่สามารถเชื่อมต่อ PeerJS โดยตรงได้ภายในเวลาที่กำหนด เครือข่ายอาจมีไฟร์วอลล์หรือต้องตั้งค่า TURN Server กรุณาลองสร้างห้องใหม่หรือสลับเครือข่าย",
          true
        );
      }
    }, ONLINE_CONFIG.PEER_CONNECT_TIMEOUT_MS);
  }

  setupDataConnection(connection) {
    if (this.conn && this.conn !== connection) {
      try {
        this.conn.close();
      } catch (e) {}
    }

    this.conn = connection;

    connection.on('open', () => {
      clearTimeout(this.peerConnectTimer);
      this.channelReady = true;
      this.startPingLoop();

      if (this.opponentDisconnected) {
        this.handleOpponentReconnected();
      } else {
        const oppName = this.role === 'host' ? this.clientUsername : this.hostUsername;
        if (this.role === 'host') {
          this.announcePolite(`ผู้เล่น ${oppName} เข้าร่วมครบ 2 คนและการเชื่อมต่อพร้อมแล้ว กดปุ่มเริ่มการแข่งขันออนไลน์ได้เลย`);
          this.ui.startOnlineMatchBtn.classList.remove('hidden');
        } else {
          this.announcePolite(`เชื่อมต่อกับ ${oppName} สำเร็จ รอหัวหน้าห้องกดเริ่มการแข่งขัน`);
        }
      }
    });

    connection.on('data', (data) => {
      this.handlePeerData(data);
    });

    connection.on('close', () => {
      this.channelReady = false;
      if (this.inMatch && this.matchStatus !== 'finished' && this.matchStatus !== 'aborted') {
        this.handleOpponentDisconnect();
      }
    });

    connection.on('error', (err) => {
      console.warn("DataChannel error:", err);
    });
  }

  sendPeerMessage(payload) {
    if (this.conn && this.channelReady && this.conn.open) {
      try {
        this.conn.send(payload);
        return true;
      } catch (e) {
        console.warn("Send peer message failed:", e);
      }
    }
    return false;
  }

  startPingLoop() {
    clearInterval(this.pingTimer);
    this.pingTimer = setInterval(() => {
      if (this.channelReady) {
        this.sendPeerMessage({
          type: 'ping',
          sentAt: performance.now()
        });
      }
    }, ONLINE_CONFIG.PING_INTERVAL_MS);
  }

  /**
   * ตรวจสอบสถานะห้องและเริ่มการแข่งขันโดย Host
   */
  async hostStartMatch() {
    if (this.role !== 'host' || !this.currentRoomId) return;

    const roomRef = ref(this.db, `${GAME_ID}/onlineRooms/${this.currentRoomId}`);
    const snap = await get(roomRef);
    if (!snap.exists()) {
      this.announcePolite("ไม่พบข้อมูลห้อง");
      return;
    }

    const room = snap.val();
    if (!room.host || !room.client || room.playerCount !== 2) {
      this.announcePolite("ยังไม่สามารถเริ่มได้ ต้องรอผู้เล่นครบ 2 คนก่อน");
      return;
    }

    if (!this.channelReady || !this.conn || !this.conn.open) {
      this.announcePolite("ยังไม่สามารถเริ่มได้ การเชื่อมต่อ PeerJS ยังไม่พร้อม");
      return;
    }

    this.hostSetsWon = 0;
    this.clientSetsWon = 0;
    this.currentSetNumber = 1;
    this.hostScore = 0;
    this.clientScore = 0;
    this.initialServer = 'host';
    this.currentServer = 'host';
    this.matchStatus = 'playing';

    await update(roomRef, {
      status: 'playing',
      startedAt: Date.now(),
      updatedAt: Date.now(),
      matchState: {
        hostSetsWon: 0,
        clientSetsWon: 0,
        currentSetNumber: 1,
        hostScore: 0,
        clientScore: 0,
        initialServer: 'host',
        currentServer: 'host'
      }
    });

    this.enterMatchScreenUI();

    this.sendPeerMessage({
      type: 'match_start',
      matchId: this.matchId,
      hostUsername: this.hostUsername,
      clientUsername: this.clientUsername,
      currentSetNumber: 1,
      initialServer: 'host'
    });

    this.startOnlineSet(1, 'host');
  }

  enterMatchScreenUI() {
    this.inMatch = true;
    this.ui.onlineLobbyScreen.classList.add('hidden');
    this.ui.backFromLobbyBtn.classList.add('hidden');
    this.ui.startOnlineMatchBtn.classList.add('hidden');

    this.ui.gameScreen.classList.remove('hidden');
    this.ui.leaveRoomBtn.classList.remove('hidden');

    this.motion.start();
    this.activateFocusMode();
    this.updateOnlineScoreboardUI();
  }

  activateFocusMode() {
    this.ui.gameScreen.setAttribute('role', 'application');
    this.ui.gameScreen.setAttribute('tabindex', '-1');
    try {
      this.ui.gameScreen.focus();
    } catch (e) {}
  }

  deactivateFocusMode() {
    this.ui.gameScreen.removeAttribute('role');
    this.ui.gameScreen.removeAttribute('tabindex');
  }

  getLocalRoleLabel() {
    return this.role === 'host' ? this.hostUsername : this.clientUsername;
  }

  getOpponentRoleLabel() {
    return this.role === 'host' ? this.clientUsername : this.hostUsername;
  }

  updateOnlineScoreboardUI() {
    const myName = this.getLocalRoleLabel();
    const oppName = this.getOpponentRoleLabel();
    const mySets = this.role === 'host' ? this.hostSetsWon : this.clientSetsWon;
    const oppSets = this.role === 'host' ? this.clientSetsWon : this.hostSetsWon;
    const myPts = this.role === 'host' ? this.hostScore : this.clientScore;
    const oppPts = this.role === 'host' ? this.clientScore : this.hostScore;

    const newTournamentText = `แมตช์ออนไลน์ 1 ต่อ 1 (${this.currentRoomId}): ${myName} พบกับ ${oppName}`;
    const newSetText = `เซตที่ ${this.currentSetNumber} (เซต: ${myName} ${mySets} - ${oppSets} ${oppName})`;
    const newScoreText = `แต้ม: ${myName} ${myPts} ต่อ ${oppPts} ${oppName}`;

    if (this.ui.tournamentDisplay && this.ui.tournamentDisplay.textContent !== newTournamentText) {
      this.ui.tournamentDisplay.textContent = newTournamentText;
    }
    if (this.ui.setDisplay && this.ui.setDisplay.textContent !== newSetText) {
      this.ui.setDisplay.textContent = newSetText;
    }
    if (this.ui.scoreDisplay && this.ui.scoreDisplay.textContent !== newScoreText) {
      this.ui.scoreDisplay.textContent = newScoreText;
    }
  }

  /**
   * คำนวณผู้เสิร์ฟตามกติกาปิงปองมาตรฐาน (สลับเสิร์ฟทุก 2 แต้ม และสลับทุก 1 แต้มเมื่อดิวซ์ 10-10)
   */
  computeServerForCurrentScore() {
    const totalPoints = this.hostScore + this.clientScore;
    const oppositeServer = this.initialServer === 'host' ? 'client' : 'host';

    if (this.hostScore >= 10 && this.clientScore >= 10) {
      const deucePoints = totalPoints - 20;
      return (deucePoints % 2 === 0) ? this.initialServer : oppositeServer;
    } else {
      const pairIndex = Math.floor(totalPoints / 2);
      return (pairIndex % 2 === 0) ? this.initialServer : oppositeServer;
    }
  }

  checkOnlineSetWinner() {
    if (this.hostScore >= 11 && this.hostScore - this.clientScore >= 2) return 'host';
    if (this.clientScore >= 11 && this.clientScore - this.hostScore >= 2) return 'client';
    return null;
  }

  startOnlineSet(setNum, initialSrv) {
    this.currentSetNumber = setNum;
    this.initialServer = initialSrv;
    this.hostScore = 0;
    this.clientScore = 0;
    this.matchStatus = 'playing';
    this.updateOnlineScoreboardUI();
    this.activateFocusMode();

    this.sound.play('go', 0, 1.0, false);

    const serverName = initialSrv === 'host' ? this.hostUsername : this.clientUsername;
    const isMyServe = initialSrv === this.role;

    const msg = isMyServe
      ? `เริ่มเซตที่ ${setNum} คุณ (${serverName}) ได้เสิร์ฟก่อน เตรียมเสิร์ฟ`
      : `เริ่มเซตที่ ${setNum} ${serverName} ได้เสิร์ฟก่อน เตรียมรับลูก`;
    this.announcePolite(msg, true);

    if (this.role === 'host') {
      clearTimeout(this.activeRallyTimer);
      this.activeRallyTimer = setTimeout(() => {
        this.hostPrepareRally();
      }, 2200);
    }
  }

  /**
   * Host เริ่มต้น Rally ใหม่และส่งสัญญาณซิงก์ให้ Client
   */
  hostPrepareRally() {
    if (this.role !== 'host' || this.matchStatus !== 'playing' || this.opponentDisconnected) return;

    // ตรวจสอบว่าค่า Ping สูงเกินเกณฑ์ปลอดภัยหรือไม่
    if (this.rttMs > ONLINE_CONFIG.HIGH_PING_PAUSE_THRESHOLD_MS) {
      const warnMsg = `สัญญาณอินเทอร์เน็ตหน่วงสูงผิดปกติ (${Math.round(this.rttMs)} มิลลิวินาที) พักรอสัญญาณสักครู่เพื่อความเป็นธรรม`;
      this.announcePolite(warnMsg, true);
      this.sendPeerMessage({
        type: 'high_latency_wait',
        rttMs: Math.round(this.rttMs)
      });
      this.activeRallyTimer = setTimeout(() => this.hostPrepareRally(), 2000);
      return;
    }

    this.rallyId++;
    this.shotSeq = 0;
    this.rallyCount = 0;
    this.hostPendingReceive = null;
    this.currentServer = this.computeServerForCurrentScore();

    const payload = {
      type: 'rally_ready',
      rallyId: this.rallyId,
      hostScore: this.hostScore,
      clientScore: this.clientScore,
      currentSetNumber: this.currentSetNumber,
      server: this.currentServer
    };

    this.sendPeerMessage(payload);
    this.applyRallyReadyState(payload);
  }

  applyRallyReadyState(data) {
    clearTimeout(this.activeRallyTimer);
    this.rallyId = data.rallyId;
    this.shotSeq = 0;
    this.rallyCount = 0;
    this.hostScore = data.hostScore;
    this.clientScore = data.clientScore;
    this.currentSetNumber = data.currentSetNumber;
    this.currentServer = data.server;
    this.canLocalPlayerHit = false;

    this.updateOnlineScoreboardUI();

    const myPts = this.role === 'host' ? this.hostScore : this.clientScore;
    const oppPts = this.role === 'host' ? this.clientScore : this.hostScore;
    const serverName = this.currentServer === 'host' ? this.hostUsername : this.clientUsername;

    if (this.currentServer === this.role) {
      this.sound.play('go1', 0, 1.0, false);
      this.expectedLocalDir = 'center';
      this.canLocalPlayerHit = true;
      this.announcePolite(`แต้ม ${myPts} ต่อ ${oppPts} ตาคุณเสิร์ฟ เริ่มเลย`, true);

      if (this.role === 'host') {
        this.activeRallyTimer = setTimeout(() => {
          if (this.matchStatus === 'playing' && this.rallyCount === 0 && this.currentServer === 'host') {
            this.canLocalPlayerHit = false;
            this.sound.play('retry', 0, 1.0, false);
            this.hostResolvePoint('client', `${this.hostUsername} เสิร์ฟไม่ทันเวลา ${this.clientUsername} ได้แต้ม`, 0, 'center', true, null);
          }
        }, ONLINE_CONFIG.SERVE_TIMEOUT_MS);
      }
    } else {
      this.canLocalPlayerHit = false;
      // เล่นเสียง bot.mp3 เพื่อแจ้งให้อุปกรณ์ฝั่งรับทราบว่าถึงรอบที่ฝั่งตรงข้ามกำลังจะเสิร์ฟลูกมา
      this.sound.play('bot', 0, 1.0, false);
      this.announcePolite(`แต้ม ${myPts} ต่อ ${oppPts} รอบ ${serverName} เสิร์ฟ เตรียมตีโต้`, true);

      if (this.role === 'host' && this.currentServer === 'client') {
        this.activeRallyTimer = setTimeout(() => {
          if (this.matchStatus === 'playing' && this.rallyCount === 0 && this.currentServer === 'client') {
            this.sound.play('retry', 0, 1.0, false);
            this.hostResolvePoint('host', `${this.clientUsername} เสิร์ฟไม่ทันเวลา ${this.hostUsername} ได้แต้ม`, 0, 'center', true, null);
          }
        }, ONLINE_CONFIG.SERVE_TIMEOUT_MS + Math.min(this.rttMs, 400));
      }
    }
  }

  /**
   * รับอินพุตการตีลูกของผู้เล่นในเครื่อง (จากลูกศรคีย์บอร์ดหรือการสะบัดมือถือ)
   */
  handleLocalPlayerSwing(swingDir, powerLevel = 'normal') {
    if (!this.inMatch || this.matchStatus !== 'playing' || !this.canLocalPlayerHit || this.opponentDisconnected) {
      return false;
    }

    this.canLocalPlayerHit = false;
    clearTimeout(this.activeRallyTimer);

    const isServe = (this.rallyCount === 0 && this.currentServer === this.role);
    const reactionMs = isServe ? 0 : Math.max(0, performance.now() - this.localBouncePerfTime);

    // ตรวจสอบว่าตีตรงทิศทางลูกหรือไม่ (กรณีรับลูก)
    if (!isServe && swingDir !== this.expectedLocalDir) {
      const missPan = this.directions[this.expectedLocalDir].pan;

      if (this.role === 'host') {
        // 3. โหมดออนไลน์: ฝั่ง Host ตีวืด/ตีผิดทิศทาง ให้เล่น 3.mp3 (miss) ทุกอุปกรณ์
        this.sound.play('miss', missPan, 1.0, false);
        this.hostResolvePoint(
          'client',
          `${this.hostUsername} ตีผิดทิศทาง ${this.clientUsername} ได้แต้ม`,
          this.rallyId,
          this.expectedLocalDir,
          false,
          'host'
        );
      } else {
        // 4. โหมดออนไลน์: ฝั่ง Client ตีวืด/ตีผิดทิศทาง ให้เล่น 4.mp3 (miss4) ทุกอุปกรณ์
        this.sound.play('miss4', missPan, 1.0, false);
        this.sendPeerMessage({
          type: 'client_swing',
          rallyId: this.rallyId,
          shotSeq: this.shotSeq,
          swingDir: swingDir,
          powerLevel: powerLevel,
          reactionMs: Math.round(reactionMs),
          isMiss: true
        });
      }
      return true;
    }

    // คำนวณทิศทางเป้าหมายจากจังหวะการตีเหมือนโหมดต้นฉบับ
    let outgoingLocalDir = swingDir;
    if (!isServe) {
      const ratio = reactionMs / ONLINE_CONFIG.BASE_HIT_WINDOW_MS;
      if (ratio < 0.38) {
        outgoingLocalDir = swingDir === 'left' ? 'right' : (swingDir === 'right' ? 'left' : (Math.random() < 0.5 ? 'left' : 'right'));
      } else if (ratio > 0.78) {
        outgoingLocalDir = this.expectedLocalDir;
      } else {
        outgoingLocalDir = swingDir;
      }
    }

    if (this.role === 'host') {
      this.hostExecuteAuthoritativeShot('host', outgoingLocalDir, powerLevel, swingDir);
    } else {
      // เล่นเสียงไม้กระทบลูกทันทีฝั่ง Client เพื่อความฉับไว (Zero-Lag Local Feedback)
      this.playLocalStrikeAudio(swingDir, outgoingLocalDir, powerLevel);
      this.sendPeerMessage({
        type: 'client_swing',
        rallyId: this.rallyId,
        shotSeq: this.shotSeq,
        swingDir: swingDir,
        outgoingLocalDir: outgoingLocalDir,
        powerLevel: powerLevel,
        reactionMs: Math.round(reactionMs),
        isMiss: false
      });
    }
    return true;
  }

  playLocalStrikeAudio(originLocalDir, targetLocalDir, powerLevel) {
    const startPan = this.directions[originLocalDir].pan;
    const endPan = this.directions[targetLocalDir].pan;
    const flightTime = this.flightDurations[powerLevel] || 700;

    let hitRate = 1.0;
    let hitVol = 1.0;
    if (powerLevel === 'hard') {
      hitRate = 1.25;
      hitVol = 1.2;
    } else if (powerLevel === 'soft') {
      hitRate = 0.85;
      hitVol = 0.8;
    }

    this.sound.play('hit', startPan, hitVol, false, hitRate);
    this.sound.playBallFlight(startPan, endPan, flightTime, true, powerLevel);

    // เสียงลูกตกกระทบโต๊ะฝั่งตรงข้าม (Distant Bounce)
    setTimeout(() => {
      if (this.inMatch && this.matchStatus === 'playing') {
        this.sound.play('bounce', endPan, 0.22, true);
      }
    }, flightTime);
  }

  /**
   * Host ประมวลผลการตีที่ถูกต้องและส่งเหตุการณ์ลูกพุ่งไปยังผู้รับพร้อมคำนวณเวลาชดเชยความหน่วง
   */
  hostExecuteAuthoritativeShot(shooterRole, shooterTargetLocalDir, powerLevel, shooterOriginLocalDir) {
    if (this.role !== 'host' || this.matchStatus !== 'playing') return;
    clearTimeout(this.activeRallyTimer);
    this.hostPendingReceive = null;

    this.rallyCount++;
    this.shotSeq++;

    const currentRallyId = this.rallyId;
    const currentShotSeq = this.shotSeq;
    const receiverRole = shooterRole === 'host' ? 'client' : 'host';

    // แปลงพิกัดทิศทางจากมุมมองคนตีเป็นมุมมองคนรับฝั่งตรงข้ามโต๊ะ
    const receiverOriginDir = mirrorDirection(shooterOriginLocalDir);
    const receiverTargetDir = mirrorDirection(shooterTargetLocalDir);
    const flightTime = this.flightDurations[powerLevel] || 700;

    // หาก Host เป็นคนตี ให้เล่นเสียงตีฝั่ง Host ทันที
    if (shooterRole === 'host') {
      this.playLocalStrikeAudio(shooterOriginLocalDir, shooterTargetLocalDir, powerLevel);
    }

    const shotEvent = {
      type: 'shot_event',
      rallyId: currentRallyId,
      shotSeq: currentShotSeq,
      rallyCount: this.rallyCount,
      shooter: shooterRole,
      receiver: receiverRole,
      shooterOriginDir: shooterOriginLocalDir,
      shooterTargetDir: shooterTargetLocalDir,
      receiverOriginDir: receiverOriginDir,
      receiverTargetDir: receiverTargetDir,
      powerLevel: powerLevel,
      flightTime: flightTime,
      hitWindowMs: ONLINE_CONFIG.BASE_HIT_WINDOW_MS
    };

    this.sendPeerMessage(shotEvent);

    if (receiverRole === 'host') {
      // กรณีลูกพุ่งมาหา Host: เล่นเสียงลูกพุ่งเข้าและเปิดหน้าต่างให้ Host ตีเมื่อลูกตกกระทบโต๊ะ
      this.scheduleIncomingBallForLocalReceiver(shotEvent);
    } else {
      // กรณีลูกพุ่งไปหา Client: ตั้งเวลาตรวจสอบผลการรับลูกของ Client โดยชดเชยค่า RTT + ระยะเวลาเดินทางของลูก
      const oneWayLatency = Math.min(ONLINE_CONFIG.MAX_LATENCY_COMPENSATION_MS, Math.max(20, this.rttMs / 2));
      const hostWaitDeadlineMs = flightTime + ONLINE_CONFIG.BASE_HIT_WINDOW_MS + (oneWayLatency * 2) + 180;

      this.hostPendingReceive = {
        rallyId: currentRallyId,
        shotSeq: currentShotSeq,
        expectedClientDir: receiverTargetDir,
        maxReactionMs: ONLINE_CONFIG.BASE_HIT_WINDOW_MS + Math.min(220, oneWayLatency)
      };

      this.activeRallyTimer = setTimeout(() => {
        if (
          this.matchStatus === 'playing' &&
          this.rallyId === currentRallyId &&
          this.shotSeq === currentShotSeq &&
          this.hostPendingReceive
        ) {
          this.hostPendingReceive = null;
          // 4. โหมดออนไลน์: กรณี Client ตีไม่ทันจนหมดเวลาฝั่ง Host ให้เล่น 4.mp3 (miss4)
          this.sound.play('miss4', this.directions[shooterTargetLocalDir].pan, 0.38, true);
          this.hostResolvePoint(
            'host',
            `${this.clientUsername} ตีไม่ทัน ${this.hostUsername} ได้แต้ม`,
            currentRallyId,
            shooterTargetLocalDir,
            false,
            'client'
          );
        }
      }, hostWaitDeadlineMs);
    }
  }

  /**
   * เล่นเสียงลูกปิงปองที่พุ่งเข้ามาหาผู้รับ และเริ่มนับเวลาตีหลังจากลูกตกโต๊ะผู้รับจริงเท่านั้น
   */
  scheduleIncomingBallForLocalReceiver(shotEvent) {
    const startPan = this.directions[shotEvent.receiverOriginDir].pan;
    const endPan = this.directions[shotEvent.receiverTargetDir].pan;
    const flightTime = shotEvent.flightTime;
    const powerLevel = shotEvent.powerLevel;

    let oppHitRate = 1.0;
    if (powerLevel === 'hard') oppHitRate = 1.25;
    else if (powerLevel === 'soft') oppHitRate = 0.85;

    // เสียงไม้ฝั่งตรงข้ามตีลูกและเสียงลูกวิ่งข้ามโต๊ะมาหาเรา
    this.sound.play('hit', startPan, 0.34, true, oppHitRate);
    this.sound.playBallFlight(startPan, endPan, flightTime, false, powerLevel);

    clearTimeout(this.activeRallyTimer);
    this.activeRallyTimer = setTimeout(() => {
      if (!this.inMatch || this.matchStatus !== 'playing' || this.rallyId !== shotEvent.rallyId) return;

      // ลูกตกกระทบโต๊ะฝั่งเรา (เสียงดังชัดเจนตามตำแหน่งซ้าย/กลาง/ขวาจริง)
      const bounceVol = shotEvent.receiverTargetDir === 'center' ? 1.30 : 1.0;
      this.sound.play('bounce', endPan, bounceVol, false);

      this.rallyCount = shotEvent.rallyCount;
      this.shotSeq = shotEvent.shotSeq;
      this.expectedLocalDir = shotEvent.receiverTargetDir;
      this.localBouncePerfTime = performance.now();
      this.canLocalPlayerHit = true;

      // เริ่มจับเวลาตีหลังจากที่เครื่องผู้รับได้ยินเสียงลูกตกกระทบโต๊ะแล้วเท่านั้น
      this.activeRallyTimer = setTimeout(() => {
        if (this.canLocalPlayerHit && this.matchStatus === 'playing' && this.rallyId === shotEvent.rallyId) {
          this.canLocalPlayerHit = false;

          if (this.role === 'host') {
            // 3. โหมดออนไลน์: ฝั่ง Host ตีไม่ทัน ให้เล่น 3.mp3 (miss) ทุกอุปกรณ์
            this.sound.play('miss', endPan, 1.0, false);
            this.hostResolvePoint(
              'client',
              `${this.hostUsername} ตีไม่ทัน ${this.clientUsername} ได้แต้ม`,
              shotEvent.rallyId,
              shotEvent.receiverTargetDir,
              false,
              'host'
            );
          } else {
            // 4. โหมดออนไลน์: ฝั่ง Client ตีไม่ทัน ให้เล่น 4.mp3 (miss4) ทุกอุปกรณ์
            this.sound.play('miss4', endPan, 1.0, false);
            this.sendPeerMessage({
              type: 'client_timeout_miss',
              rallyId: shotEvent.rallyId,
              shotSeq: shotEvent.shotSeq
            });
          }
        }
      }, shotEvent.hitWindowMs);
    }, flightTime);
  }

  /**
   * รับและจัดการข้อความผ่าน PeerJS DataChannel
   */
  handlePeerData(data) {
    if (!data || !data.type) return;

    switch (data.type) {
      case 'ping': {
        this.sendPeerMessage({
          type: 'pong',
          sentAt: data.sentAt
        });
        break;
      }

      case 'pong': {
        if (typeof data.sentAt === 'number') {
          const sampleRtt = Math.max(5, performance.now() - data.sentAt);
          this.rttMs = (this.rttMs * 0.7) + (sampleRtt * 0.3);
          if (!this.inMatch && this.channelReady && this.ui.roomConnectionDisplay) {
            this.ui.roomConnectionDisplay.textContent = `สถานะห้อง: เชื่อมต่อพร้อมแข่งขัน (Ping ${Math.round(this.rttMs)} ms)`;
          }
        }
        break;
      }

      case 'match_start': {
        if (this.role === 'client') {
          this.matchId = data.matchId || this.matchId;
          this.hostUsername = data.hostUsername || this.hostUsername;
          this.clientUsername = data.clientUsername || this.clientUsername;
          this.hostSetsWon = 0;
          this.clientSetsWon = 0;
          this.enterMatchScreenUI();
          this.startOnlineSet(data.currentSetNumber || 1, data.initialServer || 'host');
        }
        break;
      }

      case 'rally_ready': {
        if (this.role === 'client') {
          this.applyRallyReadyState(data);
        }
        break;
      }

      case 'shot_event': {
        if (this.role === 'client' && data.receiver === 'client') {
          this.scheduleIncomingBallForLocalReceiver(data);
        }
        break;
      }

      case 'client_swing': {
        if (this.role === 'host') {
          this.hostProcessClientSwing(data);
        }
        break;
      }

      case 'client_timeout_miss': {
        if (this.role === 'host' && data.rallyId === this.rallyId && this.hostPendingReceive) {
          const pending = this.hostPendingReceive;
          this.hostPendingReceive = null;
          clearTimeout(this.activeRallyTimer);
          const hostSideDir = mirrorDirection(pending.expectedClientDir);
          // 4. โหมดออนไลน์: เล่น 4.mp3 (miss4) บนเครื่อง Host เมื่อ Client ตีไม่ทัน
          this.sound.play('miss4', this.directions[hostSideDir].pan, 0.38, true);
          this.hostResolvePoint(
            'host',
            `${this.clientUsername} ตีไม่ทัน ${this.hostUsername} ได้แต้ม`,
            data.rallyId,
            hostSideDir,
            false,
            'client'
          );
        }
        break;
      }

      case 'point_result': {
        if (this.role === 'client') {
          this.clientApplyPointResult(data);
        }
        break;
      }

      case 'next_set_start': {
        if (this.role === 'client') {
          this.hostSetsWon = data.hostSetsWon;
          this.clientSetsWon = data.clientSetsWon;
          this.startOnlineSet(data.currentSetNumber, data.initialServer);
        }
        break;
      }

      case 'high_latency_wait': {
        this.announcePolite(`เครือข่ายมีความหน่วงสูง (${data.rttMs} ms) กำลังปรับจังหวะรอสัญญาณ`, true);
        break;
      }

      case 'state_recovery_sync': {
        if (this.role === 'client') {
          this.applyStateRecoverySync(data);
        }
        break;
      }

      default:
        break;
    }
  }

  /**
   * Host ตรวจสอบคำสั่งตีของ Client โดยคำนึงถึง Latency Compensation
   */
  hostProcessClientSwing(data) {
    if (this.matchStatus !== 'playing' || data.rallyId !== this.rallyId) return;

    // กรณี Client เป็นผู้เสิร์ฟลูกแรกของ Rally
    if (this.rallyCount === 0 && this.currentServer === 'client' && data.shotSeq === 0) {
      clearTimeout(this.activeRallyTimer);
      this.hostExecuteAuthoritativeShot(
        'client',
        data.outgoingLocalDir || data.swingDir,
        data.powerLevel || 'normal',
        data.swingDir
      );
      return;
    }

    // กรณี Client ตีโต้ระหว่าง Rally
    if (!this.hostPendingReceive || data.shotSeq !== this.hostPendingReceive.shotSeq) {
      return;
    }

    const pending = this.hostPendingReceive;
    this.hostPendingReceive = null;
    clearTimeout(this.activeRallyTimer);

    if (data.isMiss || data.swingDir !== pending.expectedClientDir) {
      const hostSideDir = mirrorDirection(pending.expectedClientDir);
      // 4. โหมดออนไลน์: เล่น 4.mp3 (miss4) บนเครื่อง Host เมื่อ Client ตีวืด/ผิดทิศทาง
      this.sound.play('miss4', this.directions[hostSideDir].pan, 0.38, true);
      this.hostResolvePoint(
        'host',
        `${this.clientUsername} ตีวืด ${this.hostUsername} ได้แต้ม`,
        this.rallyId,
        hostSideDir,
        false,
        'client'
      );
      return;
    }

    if (typeof data.reactionMs === 'number' && data.reactionMs > pending.maxReactionMs) {
      const hostSideDir = mirrorDirection(pending.expectedClientDir);
      // 4. โหมดออนไลน์: เล่น 4.mp3 (miss4) บนเครื่อง Host เมื่อ Client ตีช้าเกินจังหวะ
      this.sound.play('miss4', this.directions[hostSideDir].pan, 0.38, true);
      this.hostResolvePoint(
        'host',
        `${this.clientUsername} ตีช้าเกินจังหวะ ${this.hostUsername} ได้แต้ม`,
        this.rallyId,
        hostSideDir,
        false,
        'client',
        true
      );
      return;
    }

    this.hostExecuteAuthoritativeShot(
      'client',
      data.outgoingLocalDir || data.swingDir,
      data.powerLevel || 'normal',
      data.swingDir
    );
  }

  /**
   * Host ตัดสินผลแต้ม เซต และแมตช์ พร้อมซิงก์ไปยัง Client และ Firebase
   */
  async hostResolvePoint(
    pointWinner,
    reasonText,
    rallyId,
    hostPanDir = 'center',
    isServeTimeout = false,
    missedBy = null,
    playClientLateMiss = false
  ) {
    if (this.role !== 'host' || this.matchStatus !== 'playing') return;
    clearTimeout(this.activeRallyTimer);
    this.canLocalPlayerHit = false;
    this.hostPendingReceive = null;

    if (pointWinner === 'host') {
      this.hostScore++;
    } else {
      this.clientScore++;
    }

    const setWinner = this.checkOnlineSetWinner();
    let matchWinner = null;

    if (setWinner === 'host') {
      this.hostSetsWon++;
    } else if (setWinner === 'client') {
      this.clientSetsWon++;
    }

    if (this.hostSetsWon === 3) {
      matchWinner = 'host';
    } else if (this.clientSetsWon === 3) {
      matchWinner = 'client';
    }

    const resultPayload = {
      type: 'point_result',
      rallyId: rallyId,
      pointWinner: pointWinner,
      reasonText: reasonText,
      hostScore: this.hostScore,
      clientScore: this.clientScore,
      hostSetsWon: this.hostSetsWon,
      clientSetsWon: this.clientSetsWon,
      currentSetNumber: this.currentSetNumber,
      setWinner: setWinner,
      matchWinner: matchWinner,
      isServeTimeout: isServeTimeout,
      missedBy: missedBy,
      hostPanDir: hostPanDir,
      playClientLateMiss: playClientLateMiss
    };

    this.sendPeerMessage(resultPayload);
    await this.syncMatchStateToFirebase(setWinner, matchWinner);
    this.presentPointOrSetOutcome(resultPayload);
  }

  clientApplyPointResult(data) {
    clearTimeout(this.activeRallyTimer);
    this.canLocalPlayerHit = false;
    if (data.isServeTimeout) {
      this.sound.play('retry', 0, 1.0, false);
    } else if (data.missedBy === 'host') {
      // 3. โหมดออนไลน์: เล่น 3.mp3 (miss) บนเครื่อง Client เมื่อฝั่ง Host ตีวืด/ตีไม่ถูก/ตีผิดทิศทาง
      const clientSideDir = mirrorDirection(data.hostPanDir || 'center');
      const pan = this.directions[clientSideDir] ? this.directions[clientSideDir].pan : 0;
      this.sound.play('miss', pan, 0.38, true);
    } else if (data.missedBy === 'client' && data.playClientLateMiss) {
      // 4. โหมดออนไลน์: เล่น 4.mp3 (miss4) บนเครื่อง Client กรณีที่ Host ตัดสินว่า Client ตีช้าเกินจังหวะ
      const clientSideDir = mirrorDirection(data.hostPanDir || 'center');
      const pan = this.directions[clientSideDir] ? this.directions[clientSideDir].pan : 0;
      this.sound.play('miss4', pan, 1.0, false);
    }
    this.hostScore = data.hostScore;
    this.clientScore = data.clientScore;
    this.hostSetsWon = data.hostSetsWon;
    this.clientSetsWon = data.clientSetsWon;
    this.currentSetNumber = data.currentSetNumber;

    this.presentPointOrSetOutcome(data);
  }

  presentPointOrSetOutcome(data) {
    this.updateOnlineScoreboardUI();

    if (data.matchWinner) {
      this.matchStatus = 'finished';
      this.deactivateFocusMode();
      this.motion.stop();

      const didIWin = data.matchWinner === this.role;
      this.sound.play(didIWin ? 'winner' : 'loser', 0, 1.0, false);

      const winnerName = data.matchWinner === 'host' ? this.hostUsername : this.clientUsername;
      const finalMsg = `${data.reasonText} จบการแข่งขันออนไลน์! ผู้ชนะคือ ${winnerName} ชนะไปด้วยสกอร์ ${this.hostSetsWon} ต่อ ${this.clientSetsWon} เซต กดปุ่มออกจากห้องด้านล่างเพื่อกลับสู่หน้าหลัก`;
      this.announcePolite(finalMsg, true);

      this.recordMatchStatsIdempotent('completed', data.matchWinner);
      return;
    }

    if (data.setWinner) {
      this.matchStatus = 'between_sets';
      const didIWinSet = data.setWinner === this.role;
      this.sound.play(didIWinSet ? 'winset' : 'lose', 0, 1.0, false);

      const setWinnerName = data.setWinner === 'host' ? this.hostUsername : this.clientUsername;
      const setMsg = `${data.reasonText} จบเซตที่ ${this.currentSetNumber}! ${setWinnerName} ชนะเซตนี้ ${this.hostScore} ต่อ ${this.clientScore} แต้ม (รวมเซต ${this.hostUsername} ${this.hostSetsWon} ต่อ ${this.clientSetsWon} ${this.clientUsername}) กำลังเริ่มเซตที่ ${this.currentSetNumber + 1} อัตโนมัติ`;
      this.announcePolite(setMsg, true);

      if (this.role === 'host') {
        const nextSetNum = this.currentSetNumber + 1;
        const nextInitialServer = (nextSetNum % 2 === 1) ? 'host' : 'client';
        this.activeRallyTimer = setTimeout(() => {
          this.sendPeerMessage({
            type: 'next_set_start',
            currentSetNumber: nextSetNum,
            initialServer: nextInitialServer,
            hostSetsWon: this.hostSetsWon,
            clientSetsWon: this.clientSetsWon
          });
          this.startOnlineSet(nextSetNum, nextInitialServer);
        }, 4200);
      }
      return;
    }

    // แต้มปกติในเซต
    this.announcePolite(data.reasonText, true);
    if (this.role === 'host') {
      this.activeRallyTimer = setTimeout(() => {
        this.hostPrepareRally();
      }, ONLINE_CONFIG.POINT_TRANSITION_MS);
    }
  }

  async syncMatchStateToFirebase(setWinner, matchWinner) {
    if (!this.currentRoomId || this.role !== 'host') return;
    const roomRef = ref(this.db, `${GAME_ID}/onlineRooms/${this.currentRoomId}`);
    const updates = {
      updatedAt: Date.now(),
      matchState: {
        hostSetsWon: this.hostSetsWon,
        clientSetsWon: this.clientSetsWon,
        currentSetNumber: this.currentSetNumber,
        hostScore: this.hostScore,
        clientScore: this.clientScore,
        initialServer: this.initialServer,
        currentServer: this.currentServer,
        lastSetWinner: setWinner || null,
        matchWinner: matchWinner || null
      }
    };

    if (matchWinner) {
      updates.status = 'finished';
      updates.winnerRole = matchWinner;
      updates.winnerUsername = matchWinner === 'host' ? this.hostUsername : this.clientUsername;
      updates.finishedAt = Date.now();
    }

    try {
      await update(roomRef, updates);
    } catch (e) {
      console.warn("อัปเดตสถานะห้องไม่สำเร็จ:", e);
    }
  }

  /**
   * จัดการเมื่อผู้เล่นอีกฝ่ายหลุดจากห้องระหว่างแข่ง: พักเกมและจับเวลารอ 30 วินาที
   */
  handleOpponentDisconnect() {
    if (this.opponentDisconnected || this.matchStatus === 'finished' || this.matchStatus === 'aborted') return;
    this.opponentDisconnected = true;
    this.canLocalPlayerHit = false;
    clearTimeout(this.activeRallyTimer);

    const oppName = this.getOpponentRoleLabel();
    let remainingSec = Math.floor(ONLINE_CONFIG.RECONNECT_GRACE_MS / 1000);

    this.announcePolite(
      `ผู้เล่น ${oppName} การเชื่อมต่อหลุด ระบบพักการแข่งขันชั่วคราวและรอการเชื่อมต่อกลับภายใน ${remainingSec} วินาที`,
      true
    );

    clearInterval(this.reconnectCountdownInterval);
    this.reconnectCountdownInterval = setInterval(() => {
      remainingSec--;
      if (remainingSec > 0 && this.inMatch && this.ui.gameStatus) {
        this.ui.gameStatus.textContent = `รอผู้เล่น ${oppName} เชื่อมต่อกลับ... เหลือเวลา ${remainingSec} วินาที`;
      }
    }, 1000);

    clearTimeout(this.reconnectTimer);
    this.reconnectTimer = setTimeout(async () => {
      clearInterval(this.reconnectCountdownInterval);
      if (this.opponentDisconnected && this.matchStatus !== 'finished' && this.matchStatus !== 'aborted') {
        await this.abortMatchDueToDisconnect(
          `ผู้เล่น ${oppName} ไม่สามารถเชื่อมต่อกลับมาได้ภายใน 30 วินาที ยุติการแข่งขันโดยไม่นับผลแพ้ชนะปกติ`
        );
      }
    }, ONLINE_CONFIG.RECONNECT_GRACE_MS);
  }

  handleOpponentReconnected() {
    if (!this.opponentDisconnected) return;
    this.opponentDisconnected = false;
    clearTimeout(this.reconnectTimer);
    clearInterval(this.reconnectCountdownInterval);

    const oppName = this.getOpponentRoleLabel();
    this.announcePolite(`ผู้เล่น ${oppName} เชื่อมต่อกลับมาสำเร็จ ดำเนินการแข่งขันต่อจากแต้มเดิม`, true);

    if (this.role === 'host' && this.inMatch && this.matchStatus === 'playing') {
      this.sendPeerMessage({
        type: 'state_recovery_sync',
        hostSetsWon: this.hostSetsWon,
        clientSetsWon: this.clientSetsWon,
        currentSetNumber: this.currentSetNumber,
        hostScore: this.hostScore,
        clientScore: this.clientScore,
        initialServer: this.initialServer
      });
      this.activeRallyTimer = setTimeout(() => {
        this.hostPrepareRally();
      }, 2000);
    }
  }

  applyStateRecoverySync(data) {
    this.hostSetsWon = data.hostSetsWon;
    this.clientSetsWon = data.clientSetsWon;
    this.currentSetNumber = data.currentSetNumber;
    this.hostScore = data.hostScore;
    this.clientScore = data.clientScore;
    this.initialServer = data.initialServer;
    this.matchStatus = 'playing';
    this.enterMatchScreenUI();
  }

  async abortMatchDueToDisconnect(reason) {
    if (this.currentRoomId) {
      try {
        await update(ref(this.db, `${GAME_ID}/onlineRooms/${this.currentRoomId}`), {
          status: 'aborted',
          abortReason: reason,
          finishedAt: Date.now(),
          updatedAt: Date.now()
        });
      } catch (e) {}
    }
    this.finishMatchAsAborted(reason);
  }

  finishMatchAsAborted(reason) {
    if (this.matchStatus === 'aborted' || this.matchStatus === 'finished') return;
    this.matchStatus = 'aborted';
    this.canLocalPlayerHit = false;
    clearTimeout(this.activeRallyTimer);
    clearTimeout(this.reconnectTimer);
    clearInterval(this.reconnectCountdownInterval);

    this.deactivateFocusMode();
    this.motion.stop();
    this.announcePolite(reason, true);
    this.recordMatchStatsIdempotent('aborted', null);
  }

  /**
   * บันทึกสถิติออนไลน์ลง PingPong/onlineStats/ โดยรับประกันว่าบันทึกเพียงครั้งเดียวต่อ Match ID (Idempotent Transaction)
   * และไม่กระทบคะแนน Tournament หรือ Hall of Fame เดิม
   */
  async recordMatchStatsIdempotent(outcomeType, winnerRole) {
    const user = this.getCurrentUser();
    if (!user || !this.matchId || !this.currentRoomId) return;

    const myKey = encodeUserKey(user.username);
    const statsRef = ref(this.db, `${GAME_ID}/onlineStats/${myKey}`);
    const matchId = this.matchId;
    const isWin = outcomeType === 'completed' && winnerRole === this.role;
    const isLoss = outcomeType === 'completed' && winnerRole && winnerRole !== this.role;
    const isAborted = outcomeType === 'aborted';

    try {
      await runTransaction(statsRef, (currentStats) => {
        const st = currentStats || {
          username: user.username,
          totalMatches: 0,
          wins: 0,
          losses: 0,
          aborted: 0,
          lastPlayedAt: 0,
          processedMatches: {}
        };

        if (!st.processedMatches) st.processedMatches = {};
        if (st.processedMatches[matchId]) {
          // เคยบันทึก Match ID นี้ไปแล้ว ไม่เพิ่มซ้ำ
          return st;
        }

        st.processedMatches[matchId] = {
          roomId: this.currentRoomId,
          outcome: outcomeType,
          recordedAt: Date.now()
        };
        st.totalMatches = (st.totalMatches || 0) + 1;
        if (isWin) st.wins = (st.wins || 0) + 1;
        if (isLoss) st.losses = (st.losses || 0) + 1;
        if (isAborted) st.aborted = (st.aborted || 0) + 1;
        st.lastPlayedAt = Date.now();

        return st;
      });
      await this.loadPlayerOnlineStatsUI();
    } catch (e) {
      console.warn("บันทึกสถิติออนไลน์ไม่สำเร็จ:", e);
    }
  }

  updateRoomPanelVisibility(inRoom) {
    if (inRoom) {
      this.ui.activeRoomCard.classList.remove('hidden');
      this.ui.createRoomBtn.classList.add('hidden');
      this.ui.leaveRoomBtn.classList.remove('hidden');
    } else {
      this.ui.activeRoomCard.classList.add('hidden');
      this.ui.createRoomBtn.classList.remove('hidden');
      this.ui.leaveRoomBtn.classList.add('hidden');
      this.ui.startOnlineMatchBtn.classList.add('hidden');
    }
  }

  cleanupPeerOnly() {
    clearTimeout(this.peerConnectTimer);
    clearInterval(this.pingTimer);
    this.channelReady = false;
    this.peerReady = false;

    if (this.conn) {
      try {
        this.conn.close();
      } catch (e) {}
      this.conn = null;
    }

    if (this.peer) {
      try {
        this.peer.destroy();
      } catch (e) {}
      this.peer = null;
    }
  }

  /**
   * ออกจากห้องพร้อม Cleanup Firebase Listener, PeerJS Connection, Event Listener และ Timer ทั้งหมด
   */
  async leaveCurrentRoom(notifyUser = true) {
    clearTimeout(this.activeRallyTimer);
    clearTimeout(this.reconnectTimer);
    clearInterval(this.reconnectCountdownInterval);

    const leavingRoomId = this.currentRoomId;
    const leavingRole = this.role;
    const wasInMatch = this.inMatch;
    const prevStatus = this.matchStatus;

    this.currentRoomId = null;
    this.role = null;
    this.inMatch = false;
    this.opponentDisconnected = false;
    this.canLocalPlayerHit = false;
    this.lastJoinedClientKey = null;
    sessionStorage.removeItem('pingpong_active_online_room');

    if (this.roomUnsubscribe) {
      this.roomUnsubscribe();
      this.roomUnsubscribe = null;
    }

    if (this.presenceDisconnectRef) {
      try {
        await this.presenceDisconnectRef.cancel();
      } catch (e) {}
      this.presenceDisconnectRef = null;
    }

    this.cleanupPeerOnly();
    this.motion.stop();
    this.deactivateFocusMode();

    if (leavingRoomId) {
      const roomRef = ref(this.db, `${GAME_ID}/onlineRooms/${leavingRoomId}`);
      try {
        const snap = await get(roomRef);
        if (snap.exists()) {
          const room = snap.val();
          if (room.status === 'waiting' && leavingRole === 'host') {
            await remove(roomRef);
          } else if (room.status === 'ready' && !wasInMatch) {
            if (leavingRole === 'host') {
              await remove(roomRef);
            } else if (leavingRole === 'client') {
              await update(roomRef, {
                client: null,
                playerCount: 1,
                status: 'waiting',
                updatedAt: Date.now()
              });
            }
          } else if (wasInMatch && prevStatus !== 'finished' && prevStatus !== 'aborted') {
            const leaverName = this.getCurrentUser() ? this.getCurrentUser().username : 'ผู้เล่น';
            await update(roomRef, {
              status: 'aborted',
              abortReason: `${leaverName} ออกจากห้องระหว่างการแข่งขัน ยุติการแข่งขันโดยไม่นับผลชนะปกติ`,
              finishedAt: Date.now(),
              updatedAt: Date.now()
            });
          }
        }
      } catch (e) {
        console.warn("Cleanup room error:", e);
      }
    }

    this.updateRoomPanelVisibility(false);

    if (wasInMatch) {
      this.ui.gameScreen.classList.add('hidden');
      this.ui.onlineLobbyScreen.classList.remove('hidden');
      this.ui.backFromLobbyBtn.classList.remove('hidden');
      await this.loadPlayerOnlineStatsUI();
    }

    if (notifyUser) {
      this.announcePolite("ออกจากห้องออนไลน์เรียบร้อยแล้ว", true);
    }
  }
}