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
  runTransaction,
  serverTimestamp
} from "https://www.gstatic.com/firebasejs/10.8.0/firebase-database.js";
import { database, GAME_ID, encodeUserKey, SessionStorage } from "./auth.js";

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
  RECONNECT_GRACE_MS: 3000,
  PEER_CONNECT_TIMEOUT_MS: 15000,
  PING_INTERVAL_MS: 2000,
  WIN_REWARD_POINTS: 100
};

/**
 * ค่ากำหนดระบบ Presence / Heartbeat / ตรวจสอบห้องหมดอายุ (เฉพาะห้องสถานะ waiting)
 * - HEARTBEAT_INTERVAL_MS: เจ้าของห้องส่งสัญญาณว่ายังออนไลน์ทุก 20 วินาที (ไม่ถี่เกินจำเป็น)
 * - HEARTBEAT_STALE_MS: ไม่มี Heartbeat เกิน 75 วินาที (ประมาณ 3-4 รอบที่ขาดหาย) จึงถือว่าหมดอายุ ไม่ตัดสินจากความล่าช้าครั้งเดียว
 * - DISCONNECT_GRACE_MS: หลัง onDisconnect ทำงาน ให้เวลาผ่อนผัน 30 วินาทีสำหรับ Refresh / สลับแอปแล้วกลับมา
 * - LEGACY_STALE_MS: ห้องรุ่นเก่าที่ไม่มี Heartbeat ใช้เกณฑ์ปลอดภัย 5 นาที
 */
const ROOM_PRESENCE_CONFIG = {
  HEARTBEAT_INTERVAL_MS: 20000,
  HEARTBEAT_STALE_MS: 75000,
  DISCONNECT_GRACE_MS: 30000,
  LEGACY_STALE_MS: 300000,
  LEGACY_CLOCK_SKEW_TOLERANCE_MS: 600000,
  LOBBY_SWEEP_MS: 30000,
  CLEANUP_COOLDOWN_MS: 10000,
  CLOCK_WAIT_TIMEOUT_MS: 4000,
  ONDISCONNECT_RETRIES: 3
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
    this.onUserUpdated = options.onUserUpdated;
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

    // ระบบ Presence / Heartbeat / Cleanup ห้องค้าง
    this.sessionId = `s_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
    this.localPeerId = null;
    this.serverTimeOffset = 0;
    this.serverClockReady = false;
    this.serverClockWaiters = [];
    this.serverOffsetUnsub = null;
    this.presenceToken = 0;
    this.presenceActive = false;
    this.presenceAnnounced = false;
    this.presenceConnUnsub = null;
    this.presenceBusyToken = null;
    this.presenceRerun = false;
    this.presenceWakeHandler = null;
    this.firebaseConnected = false;
    this.heartbeatTimer = null;
    this.heartbeatInFlight = false;
    this.heartbeatInFlightSince = 0;
    this.lobbySweepTimer = null;
    this.lastRoomsObj = null;
    this.renderedRoomsSig = '';
    this.cleanupAttemptAt = {};

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

    // ป้องกันการประมวลผลและเล่นเสียงเหตุการณ์แต้ม/ตีพลาดซ้ำจาก PeerJS หรือ Firebase Listener
    this.pointEventSeq = 0;
    this.lastProcessedPointEventId = null;
    this.lastPlayedMissAudioEventId = null;

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

    if (this.ui.btnHomeMain) {
      this.ui.btnHomeMain.classList.remove('hidden');
    }
    this.ui.onlineLobbyScreen.classList.remove('hidden');
    this.ui.backFromLobbyBtn.classList.remove('hidden');

    await this.loadPlayerOnlineStatsUI();

    // ซิงก์เวลา Firebase Server ก่อนตรวจสอบห้องหมดอายุ เพื่อไม่ให้นาฬิกาเครื่องที่ไม่ตรงทำให้ลบห้องผิด
    this.ensureServerClockSync();
    await this.waitForServerClock(ROOM_PRESENCE_CONFIG.CLOCK_WAIT_TIMEOUT_MS);

    this.subscribeWaitingRooms();

    // ตรวจสอบว่าผู้เล่นมีห้องที่กำลังเปิดค้างไว้หรือไม่ (เช่น กรณีกด Refresh)
    await this.tryRecoverExistingRoom();

    // Auto Cleanup: ลบห้อง waiting ที่ยืนยันแล้วว่าหมดอายุ (ใช้ Transaction ตรวจซ้ำก่อนลบทุกห้อง)
    await this.runLobbyCleanupNow();

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
      this.handleRoomsUpdate(snapshot.exists() ? snapshot.val() : {});
    });

    // Timer ตรวจรอบเดียวเท่านั้น (ล้างของเดิมก่อนสร้างใหม่เสมอ) เพื่อจับห้องที่หมดอายุโดยไม่มีข้อมูลเปลี่ยนแปลง
    clearInterval(this.lobbySweepTimer);
    this.lobbySweepTimer = setInterval(() => {
      if (!this.activeMode) return;
      this.sweepLobbyRooms();
    }, ROOM_PRESENCE_CONFIG.LOBBY_SWEEP_MS);
  }

  unsubscribeWaitingRooms() {
    if (this.lobbyUnsubscribe) {
      this.lobbyUnsubscribe();
      this.lobbyUnsubscribe = null;
    }
    clearInterval(this.lobbySweepTimer);
    this.lobbySweepTimer = null;
    this.lastRoomsObj = null;
  }

  RefreshRoomsManual() {
    get(ref(this.db, `${GAME_ID}/onlineRooms`)).then((snap) => {
      this.handleRoomsUpdate(snap.exists() ? snap.val() : {});
    }).catch(() => {});
  }

  /**
   * รับข้อมูลรายการห้องล่าสุด: แสดงเฉพาะห้องที่ยังใช้งานได้ และสั่งลบห้องที่ยืนยันว่าหมดอายุออกจาก Firebase
   */
  handleRoomsUpdate(roomsObj) {
    this.lastRoomsObj = roomsObj || {};
    this.renderWaitingRoomsList(this.lastRoomsObj);
    this.cleanupExpiredRooms(this.lastRoomsObj).catch(() => {});
  }

  sweepLobbyRooms() {
    const rooms = this.lastRoomsObj;
    if (!rooms) return;
    const sig = this.getJoinableRooms(rooms).map((r) => r.roomId).sort().join('|');
    if (sig !== this.renderedRoomsSig) {
      this.renderWaitingRoomsList(rooms);
    }
    this.cleanupExpiredRooms(rooms).catch(() => {});
  }

  async runLobbyCleanupNow() {
    try {
      const snap = await get(ref(this.db, `${GAME_ID}/onlineRooms`));
      await this.cleanupExpiredRooms(snap.exists() ? snap.val() : {});
    } catch (e) {
      console.warn("ตรวจสอบห้องค้างไม่สำเร็จ:", e);
    }
  }

  // ---------- เวลา Firebase Server ----------
  ensureServerClockSync() {
    if (this.serverOffsetUnsub) return;
    try {
      this.serverOffsetUnsub = onValue(ref(this.db, '.info/serverTimeOffset'), (snap) => {
        const v = snap.val();
        if (typeof v === 'number' && isFinite(v)) {
          this.serverTimeOffset = v;
          this.serverClockReady = true;
          const waiters = this.serverClockWaiters;
          this.serverClockWaiters = [];
          waiters.forEach((fn) => fn());
        }
      });
    } catch (e) {
      this.serverOffsetUnsub = null;
    }
  }

  waitForServerClock(timeoutMs) {
    if (this.serverClockReady) return Promise.resolve(true);
    return new Promise((resolve) => {
      const timer = setTimeout(() => resolve(this.serverClockReady), timeoutMs);
      this.serverClockWaiters.push(() => {
        clearTimeout(timer);
        resolve(true);
      });
    });
  }

  getServerNow() {
    return Date.now() + (this.serverTimeOffset || 0);
  }

  // ---------- ประเมินสถานะห้อง waiting ----------
  isOrphanRoomFragment(room) {
    if (room === null || room === undefined) return false;
    if (typeof room !== 'object') return true;
    return !room.roomId && !room.status && !room.matchId;
  }

  /**
   * ประเมินห้อง waiting ด้วยเวลา Firebase Server
   * joinable = แสดงปุ่มเข้าร่วมได้ | expired = ยืนยันว่าหมดอายุและลบได้
   * ห้องที่ไม่ใช่ waiting (ready / playing / finished ...) จะไม่ถูกตัดสินว่าหมดอายุที่นี่ เพื่อไม่กระทบการแข่งขันและระบบกลับเข้าห้อง
   */
  evaluateWaitingRoom(room, now) {
    const C = ROOM_PRESENCE_CONFIG;
    const num = (v) => (typeof v === 'number' && isFinite(v) ? v : null);
    if (!room || typeof room !== 'object') return { joinable: false, expired: false };
    if (room.status !== 'waiting') return { joinable: false, expired: false };

    const host = room.host;
    if (!host || typeof host !== 'object') return { joinable: false, expired: true };
    if (room.client) return { joinable: false, expired: false };

    const hb = num(host.heartbeatAt);
    const dc = num(host.disconnectedAt);

    // ห้องที่ใช้ระบบ Presence ใหม่ (เวลาจาก Firebase Server)
    if (hb !== null) {
      const evidence = Math.max(hb, dc !== null ? dc : 0);
      const age = now - evidence;
      if (host.connected === false) {
        return { joinable: false, expired: age > C.DISCONNECT_GRACE_MS };
      }
      return { joinable: age <= C.HEARTBEAT_STALE_MS, expired: age > C.HEARTBEAT_STALE_MS };
    }

    // ห้องรุ่นเก่า (ไม่มี Heartbeat) ใช้เกณฑ์ปลอดภัยจากเวลาที่บันทึกไว้
    const candidates = [host.lastSeen, room.updatedAt, room.createdAt, host.joinedAt]
      .map(num)
      .filter((v) => v !== null);
    if (candidates.length === 0) return { joinable: false, expired: true };
    const evidence = Math.max(...candidates);
    const age = now - evidence;
    if (age < -C.LEGACY_CLOCK_SKEW_TOLERANCE_MS) {
      return { joinable: false, expired: true };
    }
    if (host.connected === false) {
      return { joinable: false, expired: age > C.DISCONNECT_GRACE_MS };
    }
    return { joinable: age <= C.LEGACY_STALE_MS, expired: age > C.LEGACY_STALE_MS };
  }

  getJoinableRooms(roomsObj) {
    const now = this.getServerNow();
    return Object.values(roomsObj || {}).filter((r) => {
      return r && typeof r === 'object' && r.roomId && r.status === 'waiting' && r.host && !r.client &&
        this.evaluateWaitingRoom(r, now).joinable;
    });
  }

  /**
   * ลบห้อง waiting ที่หมดอายุ (และเศษข้อมูลห้องที่ไม่สมบูรณ์) ออกจาก Firebase
   * - ข้ามห้องของตัวเองที่กำลังใช้งาน
   * - ลบผ่าน Transaction ที่ตรวจเงื่อนไขซ้ำกับข้อมูลล่าสุดบน Server ก่อนลบทุกครั้ง
   * - ไม่แตะข้อมูลผู้เล่น สถิติ หรือห้องที่ไม่ใช่ waiting
   */
  async cleanupExpiredRooms(roomsObj) {
    if (!this.serverClockReady) return;
    const now = this.getServerNow();
    const targets = [];
    Object.keys(roomsObj || {}).forEach((key) => {
      if (key === this.currentRoomId) return;
      const last = this.cleanupAttemptAt[key] || 0;
      if (Date.now() - last < ROOM_PRESENCE_CONFIG.CLEANUP_COOLDOWN_MS) return;
      const room = roomsObj[key];
      if (this.isOrphanRoomFragment(room) || this.evaluateWaitingRoom(room, now).expired) {
        targets.push(key);
      }
    });
    for (const key of targets) {
      this.cleanupAttemptAt[key] = Date.now();
      await this.removeRoomIfExpired(key);
    }
  }

  async removeRoomIfExpired(roomKey) {
    if (!roomKey || !this.serverClockReady) return false;
    if (roomKey === this.currentRoomId) return false;
    try {
      const res = await runTransaction(ref(this.db, `${GAME_ID}/onlineRooms/${roomKey}`), (room) => {
        if (room === null) return room;
        if (this.isOrphanRoomFragment(room)) return null;
        if (this.evaluateWaitingRoom(room, this.getServerNow()).expired) return null;
        return; // ห้องกลับมาใช้งานแล้วหรือไม่ใช่ waiting: ยกเลิกการลบ
      });
      return Boolean(res.committed && !res.snapshot.exists());
    } catch (e) {
      console.warn("ลบห้องหมดอายุไม่สำเร็จ (อาจถูกจำกัดโดย Firebase Rules):", roomKey, e);
      return false;
    }
  }

  /**
   * ลบห้องของตัวเองเมื่อออกจากห้อง โดยไม่ลบหากห้องถูกกู้คืนโดยเซสชันอื่นแล้ว
   */
  async removeRoomIfOwned(roomRef) {
    await runTransaction(roomRef, (room) => {
      if (room === null) return room;
      const hostSession = room.host && room.host.sessionId;
      if (hostSession && hostSession !== this.sessionId) return;
      return null;
    });
  }

  renderWaitingRoomsList(roomsObj) {
    const container = this.ui.waitingRoomsContainer;
    if (!container) return;
    container.innerHTML = '';

    const entries = this.getJoinableRooms(roomsObj);
    this.renderedRoomsSig = entries.map((r) => r.roomId).sort().join('|');

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

      // ห้อง waiting ที่หมดช่วงผ่อนผันแล้ว ไม่สามารถกู้คืนได้ และต้องถูกลบ
      if (
        room.status === 'waiting' &&
        room.host && room.host.userKey === myKey &&
        this.serverClockReady &&
        this.evaluateWaitingRoom(room, this.getServerNow()).expired
      ) {
        sessionStorage.removeItem('pingpong_active_online_room');
        await this.removeRoomIfExpired(savedRoomId);
        this.announcePolite(`ห้อง ${savedRoomId} หมดอายุแล้ว ไม่สามารถกู้คืนได้ กรุณาสร้างห้องใหม่`, true);
        return;
      }

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
          sessionId: this.sessionId,
          heartbeatAt: serverTimestamp(),
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
      let expiredDetected = false;

      const txResult = await runTransaction(roomRef, (room) => {
        if (room === null) {
          return room;
        }
        if (room.host && room.host.userKey === myKey) {
          rejectReason = "คุณเป็นเจ้าของห้องนี้อยู่แล้ว";
          return;
        }
        if (
          this.serverClockReady &&
          room.status === 'waiting' &&
          !room.client &&
          !this.evaluateWaitingRoom(room, this.getServerNow()).joinable
        ) {
          expiredDetected = true;
          rejectReason = `ห้อง ${roomId} หมดอายุแล้ว เพราะเจ้าของห้องไม่ได้ออนไลน์ ระบบกำลังอัปเดตรายการห้อง`;
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
          sessionId: this.sessionId,
          joinedAt: Date.now()
        };
        room.playerCount = 2;
        room.status = 'ready';
        room.updatedAt = Date.now();
        return room;
      });

      if (!txResult.committed || !txResult.snapshot.exists()) {
        this.announcePolite(rejectReason || "ไม่สามารถเข้าร่วมห้องนี้ได้ ห้องอาจเต็มหรือถูกปิดไปแล้ว");
        if (expiredDetected) {
          await this.removeRoomIfExpired(roomId);
        }
        this.RefreshRoomsManual();
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

    this.peer.on('open', (peerId) => {
      this.peerReady = true;
      this.localPeerId = peerId;
      // ลงทะเบียน onDisconnect ก่อน แล้วจึงประกาศสถานะออนไลน์ + เริ่ม Heartbeat
      this.startPresence(roomId, role);
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

  // =====================================================================
  // Presence / Heartbeat (Firebase onDisconnect + .info/connected + Heartbeat เวลา Server)
  // =====================================================================

  startPresence(roomId, role) {
    this.stopPresence();
    this.ensureServerClockSync();
    const token = ++this.presenceToken;
    this.presenceActive = true;
    this.presenceAnnounced = false;
    this.presenceRerun = false;

    // ทุกครั้งที่ Firebase เชื่อมต่อใหม่ (เช่น เน็ตกลับมา / สลับแอปกลับมา) ต้องลงทะเบียน onDisconnect และประกาศออนไลน์ใหม่
    this.presenceConnUnsub = onValue(ref(this.db, '.info/connected'), (snap) => {
      const isConn = snap.val() === true;
      this.firebaseConnected = isConn;
      if (isConn) {
        this.registerPresenceAndAnnounce(roomId, role, token);
      }
    });

    clearInterval(this.heartbeatTimer);
    this.heartbeatTimer = setInterval(() => {
      this.sendHeartbeat(roomId, role, token);
    }, ROOM_PRESENCE_CONFIG.HEARTBEAT_INTERVAL_MS);

    // Safari iOS อาจหยุด JavaScript ขณะอยู่เบื้องหลัง: ส่ง Heartbeat ทันทีเมื่อกลับมาใช้งาน
    this.presenceWakeHandler = () => {
      if (document.visibilityState === 'hidden') return;
      this.sendHeartbeat(roomId, role, token);
    };
    document.addEventListener('visibilitychange', this.presenceWakeHandler);
    window.addEventListener('pageshow', this.presenceWakeHandler);
    window.addEventListener('online', this.presenceWakeHandler);
  }

  stopPresence() {
    this.presenceToken++;
    this.presenceActive = false;
    this.presenceAnnounced = false;
    this.presenceRerun = false;
    clearInterval(this.heartbeatTimer);
    this.heartbeatTimer = null;
    if (this.presenceConnUnsub) {
      try {
        this.presenceConnUnsub();
      } catch (e) {}
      this.presenceConnUnsub = null;
    }
    if (this.presenceWakeHandler) {
      document.removeEventListener('visibilitychange', this.presenceWakeHandler);
      window.removeEventListener('pageshow', this.presenceWakeHandler);
      window.removeEventListener('online', this.presenceWakeHandler);
      this.presenceWakeHandler = null;
    }
  }

  async registerPresenceAndAnnounce(roomId, role, token) {
    const isCurrent = () => token === this.presenceToken && this.currentRoomId === roomId && this.role === role;
    if (this.presenceBusyToken === token) {
      this.presenceRerun = true;
      return;
    }
    this.presenceBusyToken = token;
    try {
      do {
        this.presenceRerun = false;
        if (!isCurrent()) break;

        const roleRef = ref(this.db, `${GAME_ID}/onlineRooms/${roomId}/${role}`);

        // 1) ลงทะเบียน onDisconnect ให้สำเร็จก่อนประกาศออนไลน์ (ลองซ้ำได้หลายครั้ง)
        let registered = false;
        for (let attempt = 1; attempt <= ROOM_PRESENCE_CONFIG.ONDISCONNECT_RETRIES && !registered && isCurrent(); attempt++) {
          try {
            const od = onDisconnect(roleRef);
            await od.update({
              connected: false,
              disconnectedAt: serverTimestamp()
            });
            this.presenceDisconnectRef = od;
            registered = true;
          } catch (e) {
            console.warn(`ลงทะเบียน onDisconnect ไม่สำเร็จ (ครั้งที่ ${attempt}):`, e);
            if (attempt < ROOM_PRESENCE_CONFIG.ONDISCONNECT_RETRIES) {
              await new Promise((r) => setTimeout(r, 1500));
            }
          }
        }
        if (!isCurrent()) break;
        if (!registered) {
          // สำรอง: ยังคงมี Heartbeat ให้ระบบตัดห้องหมดอายุได้ แม้ onDisconnect ลงทะเบียนไม่สำเร็จ และจะลองใหม่เมื่อเชื่อมต่อ Firebase ครั้งถัดไป
          console.warn("onDisconnect ไม่พร้อมใช้งาน ใช้ Heartbeat เป็นตัวตรวจสอบสำรอง");
        }

        // 2) ประกาศสถานะออนไลน์
        const result = await this.announceOnline(roomId, role);
        if (!isCurrent()) break;
        if (result === 'ok') {
          this.presenceAnnounced = true;
        } else if (result === 'lost') {
          await this.handlePresenceLost(roomId, role);
          break;
        }
      } while (this.presenceRerun && isCurrent());
    } finally {
      if (this.presenceBusyToken === token) {
        this.presenceBusyToken = null;
      }
    }
  }

  async announceOnline(roomId, role) {
    const roleRef = ref(this.db, `${GAME_ID}/onlineRooms/${roomId}/${role}`);
    try {
      const res = await runTransaction(roleRef, (node) => {
        if (node === null) return node; // ห้อง/ผู้เล่นนี้ไม่มีอยู่แล้ว: ไม่สร้างข้อมูลขึ้นมาใหม่
        const now = this.getServerNow();
        node.sessionId = this.sessionId;
        node.connected = true;
        node.heartbeatAt = now;
        node.lastSeen = now;
        if (this.localPeerId) node.peerId = this.localPeerId;
        delete node.disconnectedAt;
        return node;
      }, { applyLocally: false });
      if (!res.snapshot.exists()) return 'lost';
      return res.committed ? 'ok' : 'failed';
    } catch (e) {
      console.warn("ประกาศสถานะออนไลน์ไม่สำเร็จ:", e);
      return 'failed';
    }
  }

  async sendHeartbeat(roomId, role, token) {
    if (token !== this.presenceToken || !this.presenceAnnounced || !this.firebaseConnected) return;
    if (this.currentRoomId !== roomId || this.role !== role) return;

    const nowTs = Date.now();
    if (this.heartbeatInFlight && nowTs - this.heartbeatInFlightSince < 60000) return;
    this.heartbeatInFlight = true;
    this.heartbeatInFlightSince = nowTs;

    const roleRef = ref(this.db, `${GAME_ID}/onlineRooms/${roomId}/${role}`);
    try {
      const res = await runTransaction(roleRef, (node) => {
        if (node === null) return node;
        // เซสชันอื่น (แท็บใหม่ / กู้คืนห้อง) เข้าครองห้องแล้ว: แท็บเก่าต้องไม่เขียนทับ
        if (node.sessionId && node.sessionId !== this.sessionId) return;
        const now = this.getServerNow();
        node.sessionId = this.sessionId;
        node.connected = true;
        node.heartbeatAt = now;
        node.lastSeen = now;
        delete node.disconnectedAt;
        return node;
      }, { applyLocally: false });

      if (token !== this.presenceToken) return;
      if (!res.snapshot.exists()) {
        await this.handlePresenceLost(roomId, role);
      } else if (!res.committed) {
        await this.handleSessionTakenOver(roomId, role);
      }
    } catch (e) {
      console.warn("ส่ง Heartbeat ไม่สำเร็จ:", e);
    } finally {
      this.heartbeatInFlight = false;
    }
  }

  async handlePresenceLost(roomId, role) {
    if (this.currentRoomId !== roomId || this.role !== role) return;
    if (this.inMatch) return; // ห้ามกระทบการแข่งขันที่กำลังดำเนินอยู่
    this.announcePolite("ห้องนี้หมดอายุหรือถูกปิดไปแล้ว ไม่สามารถกู้คืนได้ กรุณาสร้างห้องใหม่หรือเลือกห้องอื่น", true);
    await this.leaveCurrentRoom(false, true);
  }

  async handleSessionTakenOver(roomId, role) {
    if (this.currentRoomId !== roomId || this.role !== role) return;
    this.announcePolite("ห้องนี้ถูกเปิดใช้งานจากหน้าต่างหรือแท็บอื่นแล้ว หน้านี้จะออกจากห้องโดยไม่กระทบห้องเดิม", true);
    await this.leaveCurrentRoom(false, true);
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

    // ซิงก์เหตุการณ์ตัดสินแต้มและเสียง 3.mp3 / 4.mp3 จาก Firebase กรณีข้อความ PeerJS ล่าช้าหรือตกหล่น
    if (
      this.inMatch &&
      this.role === 'client' &&
      room.matchState &&
      room.matchState.lastPointEvent &&
      room.matchState.lastPointEvent.pointEventId
    ) {
      this.clientApplyPointResult(room.matchState.lastPointEvent);
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
    this.pointEventSeq = 0;
    this.lastProcessedPointEventId = null;
    this.lastPlayedMissAudioEventId = null;

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
    if (this.ui.btnHomeMain) {
      this.ui.btnHomeMain.classList.add('hidden');
    }
    this.ui.onlineLobbyScreen.classList.add('hidden');
    if (this.ui.backFromLobbyBtn) {
      this.ui.backFromLobbyBtn.classList.add('hidden');
    }
    this.ui.startOnlineMatchBtn.classList.add('hidden');

    this.ui.gameScreen.classList.remove('hidden');
    if (this.ui.leaveRoomBtn) {
      this.ui.leaveRoomBtn.classList.add('hidden');
    }

    if (window.visualRenderer) {
      window.visualRenderer.setupResize();
    }

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
      if (window.visualRenderer && this.directions[swingDir]) {
        window.visualRenderer.triggerSwing(true, this.directions[swingDir].pan, powerLevel);
      }
      if (this.role === 'host') {
        // 3. โหมดออนไลน์: ฝั่ง Host ตีวืด/ตีผิดทิศทาง ให้ตัดสินแต้มและกระตุ้น 3.mp3 (miss) จากเหตุการณ์ร่วมเดียวกัน
        this.hostResolvePoint(
          'client',
          `${this.hostUsername} ตีผิดทิศทาง ${this.clientUsername} ได้แต้ม`,
          this.rallyId,
          this.expectedLocalDir,
          false,
          'host'
        );
      } else {
        // 4. โหมดออนไลน์: ฝั่ง Client ตีวืด/ตีผิดทิศทาง ส่งเหตุการณ์ให้ Host ตัดสินแต้มและซิงก์เสียง 4.mp3 (miss4) พร้อมกันทั้งสองฝั่ง
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
          // 4. โหมดออนไลน์: กรณี Client ตีไม่ทันจนหมดเวลาฝั่ง Host ให้ตัดสินแต้มและเล่น 4.mp3 (miss4) จากเหตุการณ์ร่วมเดียวกัน
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
            // 3. โหมดออนไลน์: ฝั่ง Host ตีไม่ทัน ให้ตัดสินแต้มและเล่น 3.mp3 (miss) จากเหตุการณ์ร่วมเดียวกัน
            this.hostResolvePoint(
              'client',
              `${this.hostUsername} ตีไม่ทัน ${this.clientUsername} ได้แต้ม`,
              shotEvent.rallyId,
              shotEvent.receiverTargetDir,
              false,
              'host'
            );
          } else {
            // 4. โหมดออนไลน์: ฝั่ง Client ตีไม่ทัน ส่งเหตุการณ์ให้ Host ตัดสินแต้มและซิงก์เสียง 4.mp3 (miss4) พร้อมกันทั้งสองฝั่ง
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
          this.pointEventSeq = 0;
          this.lastProcessedPointEventId = null;
          this.lastPlayedMissAudioEventId = null;
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
          // 4. โหมดออนไลน์: ตัดสินแต้มและเล่น 4.mp3 (miss4) จากเหตุการณ์ร่วมเดียวกันเมื่อ Client ตีไม่ทัน
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
      // 4. โหมดออนไลน์: ตัดสินแต้มและเล่น 4.mp3 (miss4) จากเหตุการณ์ร่วมเดียวกันเมื่อ Client ตีวืด/ผิดทิศทาง
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
      // 4. โหมดออนไลน์: ตัดสินแต้มและเล่น 4.mp3 (miss4) จากเหตุการณ์ร่วมเดียวกันเมื่อ Client ตีช้าเกินจังหวะ
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
   * เล่นเสียงเหตุการณ์จบแต้ม (3.mp3 สำหรับ Host ตีพลาด / 4.mp3 สำหรับ Client ตีพลาด / retry สำหรับเสิร์ฟไม่ทัน)
   * จากข้อมูลเหตุการณ์ร่วมเดียวกันของทั้ง Host และ Client พร้อมป้องกันการเล่นซ้ำด้วย pointEventId
   */
  playSharedPointEventAudio(data) {
    if (!data) return;
    const eventId = data.pointEventId || `${data.currentSetNumber}_${data.rallyId}_${data.hostScore}_${data.clientScore}`;
    if (this.lastPlayedMissAudioEventId === eventId) {
      return;
    }
    this.lastPlayedMissAudioEventId = eventId;

    if (data.isServeTimeout) {
      this.sound.play('retry', 0, 1.0, false);
      return;
    }

    const hostDir = (data.hostPanDir && this.directions[data.hostPanDir]) ? data.hostPanDir : 'center';
    const clientDir = mirrorDirection(hostDir);

    if (data.missedBy === 'host') {
      // 3.mp3 (miss): Host เป็นผู้ตีพลาด
      if (this.role === 'host') {
        const pan = this.directions[hostDir] ? this.directions[hostDir].pan : 0;
        this.sound.play('miss', pan, 1.0, false);
      } else {
        const pan = this.directions[clientDir] ? this.directions[clientDir].pan : 0;
        this.sound.play('miss', pan, 0.38, true);
      }
    } else if (data.missedBy === 'client') {
      // 4.mp3 (miss4): Client เป็นผู้ตีพลาด
      if (this.role === 'host') {
        const pan = this.directions[hostDir] ? this.directions[hostDir].pan : 0;
        this.sound.play('miss4', pan, 0.38, true);
      } else {
        const pan = this.directions[clientDir] ? this.directions[clientDir].pan : 0;
        this.sound.play('miss4', pan, 1.0, false);
      }
    }
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

    this.pointEventSeq++;
    const pointEventId = `${this.matchId || this.currentRoomId}_s${this.currentSetNumber}_r${rallyId}_p${this.pointEventSeq}`;

    const resultPayload = {
      type: 'point_result',
      pointEventId: pointEventId,
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

    this.lastProcessedPointEventId = pointEventId;
    this.playSharedPointEventAudio(resultPayload);
    this.sendPeerMessage(resultPayload);
    await this.syncMatchStateToFirebase(setWinner, matchWinner, resultPayload);
    this.presentPointOrSetOutcome(resultPayload);
  }

  clientApplyPointResult(data) {
    if (!data) return;
    const eventId = data.pointEventId || `${data.currentSetNumber}_${data.rallyId}_${data.hostScore}_${data.clientScore}`;
    if (this.lastProcessedPointEventId === eventId) {
      return;
    }
    this.lastProcessedPointEventId = eventId;

    clearTimeout(this.activeRallyTimer);
    this.canLocalPlayerHit = false;

    this.playSharedPointEventAudio(data);

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

      if (this.ui.btnHomeMain) {
        this.ui.btnHomeMain.classList.remove('hidden');
      }

      const didIWin = data.matchWinner === this.role;
      this.sound.play(didIWin ? 'winner' : 'loser', 0, 1.0, false);

      const winnerName = data.matchWinner === 'host' ? this.hostUsername : this.clientUsername;
      const rewardMsg = didIWin
        ? ` คุณได้รับชัยชนะสะสม +1 ครั้ง พร้อมคะแนนสะสม +${ONLINE_CONFIG.WIN_REWARD_POINTS} แต้ม!`
        : '';
      const finalMsg = `${data.reasonText} จบการแข่งขันออนไลน์! ผู้ชนะคือ ${winnerName} ชนะไปด้วยสกอร์ ${this.hostSetsWon} ต่อ ${this.clientSetsWon} เซต.${rewardMsg} กดปุ่มออกจากห้องด้านล่างเพื่อกลับสู่หน้าหลัก`;
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

  async syncMatchStateToFirebase(setWinner, matchWinner, lastPointEvent = null) {
    if (!this.currentRoomId || this.role !== 'host') return;
    const roomRef = ref(this.db, `${GAME_ID}/onlineRooms/${this.currentRoomId}`);
    const matchState = {
      hostSetsWon: this.hostSetsWon,
      clientSetsWon: this.clientSetsWon,
      currentSetNumber: this.currentSetNumber,
      hostScore: this.hostScore,
      clientScore: this.clientScore,
      initialServer: this.initialServer,
      currentServer: this.currentServer,
      lastSetWinner: setWinner || null,
      matchWinner: matchWinner || null
    };

    if (lastPointEvent) {
      matchState.lastPointEvent = {
        pointEventId: lastPointEvent.pointEventId,
        rallyId: lastPointEvent.rallyId,
        pointWinner: lastPointEvent.pointWinner,
        reasonText: lastPointEvent.reasonText,
        hostScore: lastPointEvent.hostScore,
        clientScore: lastPointEvent.clientScore,
        hostSetsWon: lastPointEvent.hostSetsWon,
        clientSetsWon: lastPointEvent.clientSetsWon,
        currentSetNumber: lastPointEvent.currentSetNumber,
        setWinner: lastPointEvent.setWinner || null,
        matchWinner: lastPointEvent.matchWinner || null,
        isServeTimeout: Boolean(lastPointEvent.isServeTimeout),
        missedBy: lastPointEvent.missedBy || null,
        hostPanDir: lastPointEvent.hostPanDir || 'center',
        playClientLateMiss: Boolean(lastPointEvent.playClientLateMiss)
      };
    }

    const updates = {
      updatedAt: Date.now(),
      matchState: matchState
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
   * จัดการเมื่อผู้เล่นอีกฝ่ายหลุดจากห้องระหว่างแข่ง: แจ้งเตือน หน่วงเวลา 3 วินาที แล้วพากลับไปที่ ../index.html
   */
  handleOpponentDisconnect() {
    if (this.opponentDisconnected || this.matchStatus === 'finished' || this.matchStatus === 'aborted') return;
    this.opponentDisconnected = true;
    this.canLocalPlayerHit = false;
    clearTimeout(this.activeRallyTimer);
    clearInterval(this.reconnectCountdownInterval);
    clearTimeout(this.reconnectTimer);

    const oppName = this.getOpponentRoleLabel();
    const disconnectMsg = `ผู้เล่น ${oppName} หลุดการเชื่อมต่อ ระบบกำลังพากลับสู่หน้าหลัก`;

    this.announcePolite(disconnectMsg, true);

    if (this.currentRoomId) {
      update(ref(this.db, `${GAME_ID}/onlineRooms/${this.currentRoomId}`), {
        status: 'aborted',
        abortReason: disconnectMsg,
        finishedAt: Date.now(),
        updatedAt: Date.now()
      }).catch(() => {});
    }

    this.reconnectTimer = setTimeout(() => {
      sessionStorage.removeItem('pingpong_active_online_room');
      window.location.href = '../index.html';
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

    setTimeout(() => {
      sessionStorage.removeItem('pingpong_active_online_room');
      window.location.href = '../index.html';
    }, ONLINE_CONFIG.RECONNECT_GRACE_MS);
  }

  /**
   * บันทึกสถิติออนไลน์ลง PingPong/onlineStats/ โดยรับประกันว่าบันทึกเพียงครั้งเดียวต่อ Match ID (Idempotent Transaction)
   * และหากผู้เล่นชนะการแข่งขัน จะเพิ่มสถิติชนะสะสม +1 ครั้ง และคะแนนสะสม +100 แต้ม ในโปรไฟล์ผู้เล่นหลักด้วย
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

    let newlyRecorded = false;

    try {
      const txResult = await runTransaction(statsRef, (currentStats) => {
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
          return;
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

      newlyRecorded = Boolean(txResult.committed);

      // หากเป็นผู้ชนะและเป็นการบันทึกผลแมตช์นี้ครั้งแรก ให้บวกสถิติชนะสะสม +1 ครั้ง และคะแนนสะสม +100 แต้มในโปรไฟล์หลักของผู้เล่น
      if (newlyRecorded && isWin) {
        const playerRef = ref(this.db, `${GAME_ID}/players/${myKey}`);
        const playerTx = await runTransaction(playerRef, (profile) => {
          if (!profile) return profile;
          profile.wins = (profile.wins || 0) + 1;
          profile.points = (profile.points || 0) + ONLINE_CONFIG.WIN_REWARD_POINTS;
          profile.updatedAt = Date.now();
          return profile;
        });

        if (playerTx.committed && playerTx.snapshot.exists()) {
          const updatedProfile = playerTx.snapshot.val();
          SessionStorage.save(updatedProfile);
          if (typeof this.onUserUpdated === 'function') {
            this.onUserUpdated(updatedProfile);
          }
        }
      }

      await this.loadPlayerOnlineStatsUI();
    } catch (e) {
      console.warn("บันทึกสถิติออนไลน์ไม่สำเร็จ:", e);
    }
  }

  updateRoomPanelVisibility(inRoom) {
    if (inRoom) {
      this.ui.activeRoomCard.classList.remove('hidden');
      this.ui.createRoomBtn.classList.add('hidden');
      if (this.ui.refreshRoomsBtn) {
        this.ui.refreshRoomsBtn.classList.add('hidden');
      }
      if (this.ui.waitingRoomsHeading) {
        this.ui.waitingRoomsHeading.classList.add('hidden');
      }
      if (this.ui.waitingRoomsContainer) {
        this.ui.waitingRoomsContainer.classList.add('hidden');
      }
      if (this.ui.leaveRoomBtn) {
        this.ui.leaveRoomBtn.classList.add('hidden');
      }
      if (this.ui.backFromLobbyBtn) {
        this.ui.backFromLobbyBtn.classList.add('hidden');
      }
    } else {
      this.ui.activeRoomCard.classList.add('hidden');
      this.ui.createRoomBtn.classList.remove('hidden');
      if (this.ui.refreshRoomsBtn) {
        this.ui.refreshRoomsBtn.classList.remove('hidden');
      }
      if (this.ui.waitingRoomsHeading) {
        this.ui.waitingRoomsHeading.classList.remove('hidden');
      }
      if (this.ui.waitingRoomsContainer) {
        this.ui.waitingRoomsContainer.classList.remove('hidden');
      }
      if (this.ui.leaveRoomBtn) {
        this.ui.leaveRoomBtn.classList.add('hidden');
      }
      if (this.ui.backFromLobbyBtn && this.activeMode && !this.inMatch) {
        this.ui.backFromLobbyBtn.classList.remove('hidden');
      }
      this.ui.startOnlineMatchBtn.classList.add('hidden');
    }
  }

  cleanupPeerOnly() {
    clearTimeout(this.peerConnectTimer);
    clearInterval(this.pingTimer);
    this.stopPresence();
    this.localPeerId = null;
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
  async leaveCurrentRoom(notifyUser = true, skipRemote = false) {
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
    this.pointEventSeq = 0;
    this.lastProcessedPointEventId = null;
    this.lastPlayedMissAudioEventId = null;
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

    if (leavingRoomId && !skipRemote) {
      const roomRef = ref(this.db, `${GAME_ID}/onlineRooms/${leavingRoomId}`);
      try {
        const snap = await get(roomRef);
        if (snap.exists()) {
          const room = snap.val();
          if (room.status === 'waiting' && leavingRole === 'host') {
            await this.removeRoomIfOwned(roomRef);
          } else if (room.status === 'ready' && !wasInMatch) {
            if (leavingRole === 'host') {
              await this.removeRoomIfOwned(roomRef);
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
      if (this.ui.btnHomeMain) {
        this.ui.btnHomeMain.classList.remove('hidden');
      }
      await this.loadPlayerOnlineStatsUI();
    }

    if (notifyUser) {
      this.announcePolite("ออกจากห้องออนไลน์เรียบร้อยแล้ว", true);
    }
  }
}