/**
 * ping-online.js
 * ระบบจัดการห้องแข่งขันออนไลน์ 1 ต่อ 1 ด้วย Firebase Realtime Database + PeerJS WebRTC DataChannel
 * รองรับ Host-Authoritative Gameplay, 3D Spatial Audio Mirroring, Latency Compensation,
 * ระบบชวนเพื่อนแข่งขันแบบ Private Room และ Screen Reader Accessibility
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
import {
  database,
  GAME_ID,
  encodeUserKey,
  SessionStorage,
  PRESENCE_STATES,
  updatePlayerPresence,
  getPlayerPresence,
  subscribeOnlinePlayers,
  verifyPlayerCanBeInvited,
  verifyCanAcceptInvitation
} from "./auth.js";

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
 * ค่ากำหนดระบบคำเชิญเพื่อน (Private Room Invitation)
 */
const INVITATION_CONFIG = {
  INVITATIONS_PATH: `${GAME_ID}/invitations`,
  EXPIRY_MS: 35000,
  PRIVATE_ROOM_JOIN_WAIT_MS: 20000
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

    // เชื่อมต่อ API จาก auth.js และ Interface จาก ping.html
    this.authManager = options.authManager || null;
    this.presenceStates = options.presenceStates || PRESENCE_STATES;
    this.updatePresenceStatusFn = options.updatePresenceStatus || updatePlayerPresence;
    this.subscribeOnlinePlayersFn = options.subscribeOnlinePlayers || subscribeOnlinePlayers;
    this.verifyPlayerCanBeInvitedFn = options.verifyPlayerCanBeInvited || verifyPlayerCanBeInvited;
    this.verifyCanAcceptInvitationFn = options.verifyCanAcceptInvitation || verifyCanAcceptInvitation;
    this.invitationBridge = options.invitationBridge || (typeof window !== 'undefined' ? window.PingPongInvitationBridge : null);

    this.db = database;
    this.activeMode = false;
    this.inMatch = false;

    this.currentRoomId = null;
    this.currentRoomVisibility = 'public'; // 'public' | 'private'
    this.role = null; // 'host' | 'client'
    this.matchId = null;
    this.latestRoomData = null;

    this.peer = null;
    this.conn = null;
    this.peerReady = false;
    this.channelReady = false;
    this.peerConnectTimer = null;
    this.connectingHostPeerId = null;

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

    // ระบบส่งและรับคำเชิญเพื่อน (Private Room Invitations)
    this.boundInboxUserKey = null;
    this.inboxUnsubscribe = null;
    this.userWatcherTimer = null;
    this.lastPresentedInvitationId = null;
    this.activeIncomingInvitation = null;
    this.outgoingInvitation = null;
    this.outgoingInviteUnsub = null;
    this.outgoingInviteDisconnectRef = null;
    this.outgoingInviteTimeoutTimer = null;
    this.outgoingTargetPresenceUnsub = null;
    this.sendingInvitationBusy = false;
    this.acceptingInvitationBusy = false;
    this.startingMatchBusy = false;
    this.inviteListFriendCountAnnounced = -1;

    this.rttMs = 80;
    this.pingTimer = null;
    this.reconnectTimer = null;
    this.reconnectCountdownInterval = null;
    this.opponentDisconnected = false;
    this.lastJoinedClientKey = null;
    this.readyAnnouncedForRoomId = null;

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
    this.initGlobalInvitationListener();
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
      if (this.startingMatchBusy || this.inMatch) return;
      await this.requestSensorAndAudioOnGesture();
      await this.hostStartMatch();
    });

    this.addAccessibleClickListener(leaveRoomBtn, async () => {
      await this.leaveCurrentRoom(true);
    });
  }

  /**
   * ตรวจสอบผู้ใช้ที่เข้าสู่ระบบและผูก Listener รับคำเชิญแบบ Real-time ตลอดเวลาที่อยู่ในเกม
   */
  initGlobalInvitationListener() {
    this.ensureServerClockSync();
    const checkAndBind = () => {
      const user = typeof this.getCurrentUser === 'function' ? this.getCurrentUser() : null;
      const nextKey = user && user.username ? encodeUserKey(user.username) : null;
      if (nextKey !== this.boundInboxUserKey) {
        this.bindIncomingInvitationInbox(nextKey);
      }
    };
    checkAndBind();
    clearInterval(this.userWatcherTimer);
    this.userWatcherTimer = setInterval(checkAndBind, 1500);
  }

  bindIncomingInvitationInbox(userKey) {
    if (this.inboxUnsubscribe) {
      try { this.inboxUnsubscribe(); } catch (e) {}
      this.inboxUnsubscribe = null;
    }
    this.boundInboxUserKey = userKey;
    this.activeIncomingInvitation = null;

    if (!userKey) return;

    const inboxRef = ref(this.db, `${INVITATION_CONFIG.INVITATIONS_PATH}/${userKey}`);
    this.inboxUnsubscribe = onValue(inboxRef, (snap) => {
      if (!snap.exists()) {
        this.handleIncomingInvitationRemoved();
        return;
      }
      this.handleIncomingInvitationSnapshot(snap.val(), userKey);
    }, (err) => {
      console.warn("Incoming invitation listener error:", err);
    });
  }

  handleIncomingInvitationRemoved() {
    if (this.activeIncomingInvitation) {
      this.activeIncomingInvitation = null;
      if (this.invitationBridge && typeof this.invitationBridge.closeInvitationModal === 'function') {
        this.invitationBridge.closeInvitationModal();
      }
      if (this.invitationBridge && typeof this.invitationBridge.resumeBotMatch === 'function') {
        this.invitationBridge.resumeBotMatch();
      }
    }
  }

  async handleIncomingInvitationSnapshot(inv, myKey) {
    if (!inv || typeof inv !== 'object') return;
    if (inv.toUserKey && inv.toUserKey !== myKey) return;

    const now = this.getServerNow();

    // यदिคำเชิญถูกยกเลิกหรือหมดอายุขณะที่หน้าต่างรับคำเชิญเปิดอยู่ ให้ปิดหน้าต่างและคืนสถานะเดิมทันที
    if (inv.status === 'cancelled' || inv.status === 'expired') {
      if (this.activeIncomingInvitation && this.activeIncomingInvitation.invitationId === inv.invitationId) {
        this.activeIncomingInvitation = null;
        if (this.invitationBridge && typeof this.invitationBridge.closeInvitationModal === 'function') {
          this.invitationBridge.closeInvitationModal();
        }
        if (this.invitationBridge && typeof this.invitationBridge.resumeBotMatch === 'function') {
          this.invitationBridge.resumeBotMatch();
        }
        this.announcePolite(`คำเชิญจาก ${inv.fromUsername || 'เพื่อน'} หมดอายุหรือถูกยกเลิกแล้ว`);
      }
      return;
    }

    if (inv.status !== 'pending') return;

    // ตรวจสอบว่าคำเชิญหมดอายุหรือยัง
    if (typeof inv.expiresAt === 'number' && now > inv.expiresAt) {
      update(ref(this.db, `${INVITATION_CONFIG.INVITATIONS_PATH}/${myKey}`), {
        status: 'expired',
        updatedAt: now
      }).catch(() => {});
      return;
    }

    // ป้องกันการรับคำเชิญหากผู้เล่นกำลังแข่งขันออนไลน์กับคนอื่น หรือมีผู้เล่นครบ 2 คนในห้องแล้ว
    if (this.inMatch || (this.currentRoomId && this.role === 'client')) {
      update(ref(this.db, `${INVITATION_CONFIG.INVITATIONS_PATH}/${myKey}`), {
        status: 'busy',
        updatedAt: now
      }).catch(() => {});
      return;
    }

    if (this.lastPresentedInvitationId === inv.invitationId) {
      return;
    }

    this.lastPresentedInvitationId = inv.invitationId;
    this.activeIncomingInvitation = inv;

    if (this.invitationBridge && typeof this.invitationBridge.showIncomingInvitation === 'function') {
      this.invitationBridge.showIncomingInvitation({
        invitationId: inv.invitationId,
        fromUsername: inv.fromUsername,
        fromUserKey: inv.fromUserKey,
        toUsername: inv.toUsername,
        toUserKey: inv.toUserKey,
        roomId: inv.roomId || null,
        createdAt: inv.createdAt,
        expiresAt: inv.expiresAt
      });
    }
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
    if (this.invitationBridge && typeof this.invitationBridge.isBotPointInProgress === 'function') {
      if (this.invitationBridge.isBotPointInProgress() && !forceDuringRally) {
        return;
      }
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

    const isInviteScreenVisible = Boolean(
      this.ui.inviteFriendsScreen && !this.ui.inviteFriendsScreen.classList.contains('hidden')
    );

    if (this.inMatch && this.ui.gameStatus) {
      if (this.ui.gameStatus.textContent === nextMsg) {
        this.ui.gameStatus.textContent = '';
        setTimeout(() => {
          if (this.ui.gameStatus) this.ui.gameStatus.textContent = nextMsg;
        }, 20);
      } else {
        this.ui.gameStatus.textContent = nextMsg;
      }
    } else if (isInviteScreenVisible && this.ui.inviteFriendsStatusBox) {
      this.ui.inviteFriendsStatusBox.classList.remove('hidden');
      if (this.ui.inviteFriendsStatusBox.textContent === nextMsg) {
        this.ui.inviteFriendsStatusBox.textContent = '';
        setTimeout(() => {
          if (this.ui.inviteFriendsStatusBox) this.ui.inviteFriendsStatusBox.textContent = nextMsg;
        }, 20);
      } else {
        this.ui.inviteFriendsStatusBox.textContent = nextMsg;
      }
    } else if (this.ui.onlineLobbyScreen && !this.ui.onlineLobbyScreen.classList.contains('hidden') && this.ui.lobbyStatusBox) {
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
      setTimeout(() => {
        if (this.ui.onlineSrAnnouncer && this.ui.onlineSrAnnouncer.textContent === nextMsg) {
          this.ui.onlineSrAnnouncer.textContent = '';
        }
      }, 1200);
    }

    clearTimeout(this.announceClearTimer);
    this.announceClearTimer = setTimeout(() => {
      this.isAnnouncing = false;
      if (this.announceQueue.length > 0) {
        this.processAnnounceQueue();
      }
    }, 450);
  }

  /**
   * ตรวจสอบจาก Snapshot ของห้องว่ามีผู้เล่นทั้ง Host และ Client ครบ 2 คนแล้วจริงหรือไม่
   */
  hasBothPlayersInRoom(room) {
    if (!room || typeof room !== 'object') return false;
    const hasHost = Boolean(room.host && (room.host.userKey || room.host.username));
    const hasClient = Boolean(
      (room.client && (room.client.userKey || room.client.username)) ||
      (this.clientUsername && this.clientUsername !== 'รอผู้เล่นเข้าร่วม...')
    );
    const countOk = typeof room.playerCount === 'number' ? (room.playerCount >= 2 || hasClient) : true;
    return hasHost && hasClient && countOk;
  }

  /**
   * อัปเดต UI สถานะห้องและปุ่มเริ่มการแข่งขันออนไลน์ให้ตรงกับสถานะจริงของทั้ง Firebase และ PeerJS
   * ป้องกัน Race Condition ไม่ว่า Firebase Snapshot หรือ PeerJS DataChannel จะพร้อมก่อนหรือหลังกัน
   */
  syncRoomReadyUIAndAnnounce(room = this.latestRoomData) {
    if (this.inMatch || !this.currentRoomId) return;

    const activeRoom = room || this.latestRoomData;
    const hasBothPlayers = this.hasBothPlayersInRoom(activeRoom);
    const isPeerChannelOpen = Boolean(this.channelReady && this.conn && this.conn.open);

    if (!hasBothPlayers) {
      this.readyAnnouncedForRoomId = null;
      if (this.ui.roomConnectionDisplay) {
        this.ui.roomConnectionDisplay.textContent = 'สถานะห้อง: รอผู้เล่นคนที่ 2 เข้าร่วม (1/2)';
      }
      if (this.ui.startOnlineMatchBtn) {
        this.ui.startOnlineMatchBtn.classList.add('hidden');
        this.ui.startOnlineMatchBtn.disabled = false;
      }
      return;
    }

    if (!isPeerChannelOpen) {
      if (this.ui.roomConnectionDisplay) {
        this.ui.roomConnectionDisplay.textContent = 'สถานะห้อง: ผู้เล่นครบ 2 คนแล้ว กำลังเชื่อมสัญญาณ PeerJS...';
      }
      if (this.ui.startOnlineMatchBtn) {
        this.ui.startOnlineMatchBtn.classList.add('hidden');
      }
      return;
    }

    // มีผู้เล่นครบ 2 คนใน Firebase และ PeerJS DataChannel เปิดพร้อมใช้งานจริงแล้วทั้งสองเงื่อนไข
    if (this.ui.roomConnectionDisplay) {
      this.ui.roomConnectionDisplay.textContent = `สถานะห้อง: เชื่อมต่อพร้อมแข่งขัน (Ping ${Math.round(this.rttMs)} ms)`;
    }

    if (this.role === 'host') {
      if (this.ui.startOnlineMatchBtn) {
        this.ui.startOnlineMatchBtn.disabled = false;
        this.ui.startOnlineMatchBtn.classList.remove('hidden');
      }
      if (this.readyAnnouncedForRoomId !== this.currentRoomId) {
        this.readyAnnouncedForRoomId = this.currentRoomId;
        const oppName = (activeRoom && activeRoom.client && activeRoom.client.username)
          ? activeRoom.client.username
          : (this.clientUsername && this.clientUsername !== 'รอผู้เล่นเข้าร่วม...' ? this.clientUsername : 'ผู้ท้าชิง');
        this.announcePolite(`ผู้เล่น ${oppName} เข้าร่วมครบ 2 คนและการเชื่อมต่อพร้อมแล้ว กดปุ่มเริ่มการแข่งขันออนไลน์ได้เลย`);
        setTimeout(() => {
          try {
            if (!this.inMatch && this.ui.startOnlineMatchBtn && !this.ui.startOnlineMatchBtn.classList.contains('hidden')) {
              this.ui.startOnlineMatchBtn.focus();
            }
          } catch (e) {}
        }, 150);
      }
    } else {
      if (this.ui.startOnlineMatchBtn) {
        this.ui.startOnlineMatchBtn.classList.add('hidden');
      }
      if (this.readyAnnouncedForRoomId !== this.currentRoomId) {
        this.readyAnnouncedForRoomId = this.currentRoomId;
        const hostName = (activeRoom && activeRoom.host && activeRoom.host.username)
          ? activeRoom.host.username
          : (this.hostUsername || 'หัวหน้าห้อง');
        this.announcePolite(`เชื่อมต่อกับ ${hostName} สำเร็จ รอหัวหน้าห้องกดเริ่มการแข่งขัน`);
      }
    }
  }

  /**
   * =====================================================================
   * งานที่ 1-6: ระบบรายชื่อเพื่อน, ส่งคำเชิญ, สร้าง/เข้า Private Room, ย้ายจากห้อง Public
   * =====================================================================
   */
  onInviteFriendsScreenOpened() {
    this.inviteListFriendCountAnnounced = -1;
    const user = this.getCurrentUser();
    const myName = user ? user.username : null;

    if (typeof this.subscribeOnlinePlayersFn === 'function') {
      const unsub = this.subscribeOnlinePlayersFn((invitableList) => {
        const count = Array.isArray(invitableList) ? invitableList.length : 0;
        if (this.inviteListFriendCountAnnounced === -1) {
          this.inviteListFriendCountAnnounced = count;
          if (count > 0) {
            this.announcePolite(`พบเพื่อนออนไลน์ที่สามารถเชิญได้ ${count} คน`);
          } else {
            this.announcePolite("ขณะนี้ไม่มีเพื่อนออนไลน์ที่ว่างให้เชิญ");
          }
        }
        try { unsub(); } catch (e) {}
      }, { onlyInvitable: true, excludeUsername: myName });
    }
  }

  onInviteFriendsScreenClosed() {
    if (this.outgoingInvitation && this.outgoingInvitation.status === 'pending') {
      this.cancelOutgoingInvitation('cancelled', true).catch(() => {});
    }
  }

  async onSelectFriendToInvite(friendPresence) {
    if (!friendPresence || !friendPresence.username) return false;
    return this.sendFriendInvitation(friendPresence.username, friendPresence);
  }

  async sendFriendInvitation(targetUsername, friendPresence = null) {
    const user = this.getCurrentUser();
    if (!user || !user.username) {
      this.announcePolite("กรุณาเข้าสู่ระบบก่อนส่งคำเชิญ");
      return false;
    }

    if (this.sendingInvitationBusy) {
      this.announcePolite("กำลังดำเนินการส่งคำเชิญ กรุณารอสักครู่");
      return false;
    }

    if (this.inMatch) {
      this.announcePolite("ไม่สามารถส่งคำเชิญระหว่างกำลังแข่งขันได้");
      return false;
    }

    if (this.outgoingInvitation && this.outgoingInvitation.status === 'pending') {
      this.announcePolite(`คุณกำลังรอคำตอบจาก ${this.outgoingInvitation.toUsername} อยู่`);
      return false;
    }

    const cleanTargetName = String(targetUsername).trim();
    const targetKey = (friendPresence && friendPresence.userKey)
      ? friendPresence.userKey
      : encodeUserKey(cleanTargetName);
    const myKey = encodeUserKey(user.username);

    if (!targetKey || targetKey === myKey) {
      this.announcePolite("ไม่สามารถเชิญตัวเองได้");
      return false;
    }

    this.sendingInvitationBusy = true;
    try {
      await this.requestSensorAndAudioOnGesture();
      this.ensureServerClockSync();
      await this.waitForServerClock(1500);

      // ตรวจสอบสถานะล่าสุดของเป้าหมายอีกครั้งก่อนส่งคำเชิญ
      const checkRes = await this.verifyPlayerCanBeInvitedFn(cleanTargetName, user.username);
      if (!checkRes || !checkRes.canInvite) {
        const reason = (checkRes && checkRes.reason) ? checkRes.reason : `ผู้เล่น ${cleanTargetName} ไม่ว่างรับคำเชิญในขณะนี้`;
        this.announcePolite(reason);
        return false;
      }

      const now = this.getServerNow();
      const invitationId = `inv_${now.toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
      const targetInboxRef = ref(this.db, `${INVITATION_CONFIG.INVITATIONS_PATH}/${targetKey}`);

      let rejectReason = "";
      const txResult = await runTransaction(targetInboxRef, (currentInv) => {
        const currentNow = this.getServerNow();
        if (currentInv && typeof currentInv === 'object') {
          const isStillPending = (currentInv.status === 'pending' || currentInv.status === 'accepted') &&
            typeof currentInv.expiresAt === 'number' &&
            currentNow < currentInv.expiresAt;
          if (isStillPending) {
            if (currentInv.fromUserKey === myKey) {
              rejectReason = `ส่งคำเชิญถึง ${cleanTargetName} ไปแล้ว กำลังรอการตอบรับ`;
            } else {
              rejectReason = `${cleanTargetName} กำลังพิจารณาคำเชิญจากผู้เล่นอื่นอยู่ กรุณาลองใหม่ภายหลัง`;
            }
            return;
          }
        }

        return {
          invitationId: invitationId,
          fromUsername: user.username,
          fromUserKey: myKey,
          fromSessionId: this.sessionId,
          toUsername: (checkRes.presence && checkRes.presence.username) ? checkRes.presence.username : cleanTargetName,
          toUserKey: targetKey,
          status: 'pending',
          roomId: null,
          visibility: 'private',
          createdAt: currentNow,
          updatedAt: currentNow,
          expiresAt: currentNow + INVITATION_CONFIG.EXPIRY_MS
        };
      }, { applyLocally: false });

      if (!txResult.committed || !txResult.snapshot.exists()) {
        this.announcePolite(rejectReason || `ไม่สามารถส่งคำเชิญให้ ${cleanTargetName} ได้ในขณะนี้`);
        return false;
      }

      const createdInv = txResult.snapshot.val();
      this.outgoingInvitation = createdInv;

      // ผูก onDisconnect เพื่อยกเลิกคำเชิญอัตโนมัติหากผู้ชวนปิดแท็บหรือหลุดการเชื่อมต่อ
      try {
        const od = onDisconnect(targetInboxRef);
        await od.update({
          status: 'cancelled',
          updatedAt: serverTimestamp()
        });
        this.outgoingInviteDisconnectRef = od;
      } catch (e) {}

      // ตั้งสถานะผู้ชวนเป็นกำลังรอคำเชิญ เพื่อป้องกันคนอื่นชวนซ้อน
      await this.updatePresenceStatusFn(this.presenceStates.INVITATION_PENDING, {
        roomId: this.currentRoomId || null
      });

      this.announcePolite(`ส่งคำเชิญถึง ${createdInv.toUsername} แล้ว กำลังรอการตอบรับ`);
      this.watchOutgoingInvitation(createdInv);
      return true;
    } catch (err) {
      console.error("Send invitation error:", err);
      this.announcePolite("เกิดข้อผิดพลาดในการส่งคำเชิญ กรุณาลองใหม่อีกครั้ง");
      return false;
    } finally {
      this.sendingInvitationBusy = false;
    }
  }

  watchOutgoingInvitation(inv) {
    this.clearOutgoingInvitationWatchers();
    const targetInboxRef = ref(this.db, `${INVITATION_CONFIG.INVITATIONS_PATH}/${inv.toUserKey}`);

    // ตั้งเวลาหมดอายุคำเชิญ
    const remainingMs = Math.max(5000, (inv.expiresAt || (this.getServerNow() + INVITATION_CONFIG.EXPIRY_MS)) - this.getServerNow());
    this.outgoingInviteTimeoutTimer = setTimeout(() => {
      if (this.outgoingInvitation && this.outgoingInvitation.invitationId === inv.invitationId && this.outgoingInvitation.status === 'pending') {
        this.handleOutgoingInvitationExpired(inv);
      }
    }, remainingMs);

    // ตรวจสอบว่าผู้รับหลุดออฟไลน์ระหว่างรอคำตอบหรือไม่
    const targetPresenceRef = ref(this.db, `${GAME_ID}/presence/${inv.toUserKey}`);
    this.outgoingTargetPresenceUnsub = onValue(targetPresenceRef, (pSnap) => {
      if (!this.outgoingInvitation || this.outgoingInvitation.invitationId !== inv.invitationId) return;
      if (this.outgoingInvitation.status !== 'pending') return;
      if (!pSnap.exists()) return;
      const pVal = pSnap.val();
      if (pVal && (pVal.online === false || pVal.connected === false || pVal.status === this.presenceStates.OFFLINE)) {
        this.cancelOutgoingInvitation('expired', false).then(() => {
          this.announcePolite(`ผู้เล่น ${inv.toUsername} ออกจากเกมหรือหลุดการเชื่อมต่อแล้ว`);
        });
      }
    }, () => {});

    // ติดตามผลการตอบรับคำเชิญแบบ Real-time
    this.outgoingInviteUnsub = onValue(targetInboxRef, async (snap) => {
      if (!this.outgoingInvitation || this.outgoingInvitation.invitationId !== inv.invitationId) return;
      if (!snap.exists()) {
        this.clearOutgoingInvitationWatchers();
        this.outgoingInvitation = null;
        this.restoreHostLobbyPresenceIfIdle();
        return;
      }

      const current = snap.val();
      if (!current || current.invitationId !== inv.invitationId) {
        this.clearOutgoingInvitationWatchers();
        this.outgoingInvitation = null;
        this.restoreHostLobbyPresenceIfIdle();
        return;
      }

      this.outgoingInvitation = current;

      if (current.status === 'declined') {
        this.clearOutgoingInvitationWatchers();
        this.outgoingInvitation = null;
        await this.cleanupInvitationNodeIfMatched(inv.toUserKey, inv.invitationId);
        this.restoreHostLobbyPresenceIfIdle();
        this.announcePolite(`${inv.toUsername} ปฏิเสธคำเชิญ`);
      } else if (current.status === 'busy') {
        this.clearOutgoingInvitationWatchers();
        this.outgoingInvitation = null;
        await this.cleanupInvitationNodeIfMatched(inv.toUserKey, inv.invitationId);
        this.restoreHostLobbyPresenceIfIdle();
        this.announcePolite(`${inv.toUsername} กำลังแข่งขันหรืออยู่ในห้องอื่นแล้ว`);
      } else if (current.status === 'expired') {
        this.clearOutgoingInvitationWatchers();
        this.outgoingInvitation = null;
        await this.cleanupInvitationNodeIfMatched(inv.toUserKey, inv.invitationId);
        this.restoreHostLobbyPresenceIfIdle();
        this.announcePolite(`${inv.toUsername} ไม่ได้ตอบรับคำเชิญภายในเวลาที่กำหนด`);
      } else if (current.status === 'accepted' && !current.roomId) {
        // เพื่อนกด "ใช่" แล้ว ให้ฝั่งผู้ชวนสร้างห้อง Private Room ทันที
        clearTimeout(this.outgoingInviteTimeoutTimer);
        await this.hostCreatePrivateRoomForAcceptedInvite(current);
      } else if (current.status === 'joined' && current.roomId) {
        // เพื่อนเข้าห้อง Private สำเร็จแล้ว ล้างโหนดคำเชิญได้
        this.clearOutgoingInvitationWatchers();
        this.outgoingInvitation = null;
        await this.cleanupInvitationNodeIfMatched(inv.toUserKey, inv.invitationId);
      }
    });
  }

  async handleOutgoingInvitationExpired(inv) {
    if (!this.outgoingInvitation || this.outgoingInvitation.invitationId !== inv.invitationId) return;
    await this.cancelOutgoingInvitation('expired', false);
    this.announcePolite(`คำเชิญถึง ${inv.toUsername} หมดอายุเนื่องจากไม่มีการตอบรับ`);
  }

  async cancelOutgoingInvitation(finalStatus = 'cancelled', silent = false) {
    const activeInv = this.outgoingInvitation;
    this.clearOutgoingInvitationWatchers();
    this.outgoingInvitation = null;

    if (activeInv && activeInv.toUserKey && activeInv.invitationId) {
      const targetInboxRef = ref(this.db, `${INVITATION_CONFIG.INVITATIONS_PATH}/${activeInv.toUserKey}`);
      try {
        await runTransaction(targetInboxRef, (node) => {
          if (!node || node.invitationId !== activeInv.invitationId) return;
          if (node.status === 'pending') {
            node.status = finalStatus;
            node.updatedAt = this.getServerNow();
            return node;
          }
          return;
        }, { applyLocally: false });
      } catch (e) {}
    }

    this.restoreHostLobbyPresenceIfIdle();
    if (!silent) {
      this.announcePolite("ยกเลิกคำเชิญเรียบร้อยแล้ว");
    }
  }

  async cleanupInvitationNodeIfMatched(targetUserKey, invitationId) {
    if (!targetUserKey || !invitationId) return;
    const targetInboxRef = ref(this.db, `${INVITATION_CONFIG.INVITATIONS_PATH}/${targetUserKey}`);
    try {
      await runTransaction(targetInboxRef, (node) => {
        if (!node || node.invitationId !== invitationId) return;
        return null;
      }, { applyLocally: false });
    } catch (e) {}
  }

  clearOutgoingInvitationWatchers() {
    clearTimeout(this.outgoingInviteTimeoutTimer);
    this.outgoingInviteTimeoutTimer = null;
    if (this.outgoingInviteUnsub) {
      try { this.outgoingInviteUnsub(); } catch (e) {}
      this.outgoingInviteUnsub = null;
    }
    if (this.outgoingTargetPresenceUnsub) {
      try { this.outgoingTargetPresenceUnsub(); } catch (e) {}
      this.outgoingTargetPresenceUnsub = null;
    }
    if (this.outgoingInviteDisconnectRef) {
      try { this.outgoingInviteDisconnectRef.cancel(); } catch (e) {}
      this.outgoingInviteDisconnectRef = null;
    }
  }

  restoreHostLobbyPresenceIfIdle() {
    if (this.inMatch) return;
    if (this.currentRoomId) {
      const hasBoth = this.hasBothPlayersInRoom(this.latestRoomData);
      this.updatePresenceStatusFn(
        hasBoth ? this.presenceStates.ONLINE_READY : this.presenceStates.WAITING_ROOM,
        {
          roomId: this.currentRoomId,
          roomRole: this.role || 'host',
          roomPlayerCount: hasBoth ? 2 : 1,
          hasClientInRoom: hasBoth,
          roomStatus: hasBoth ? 'ready' : 'waiting'
        }
      ).catch(() => {});
    } else {
      this.updatePresenceStatusFn(this.presenceStates.ONLINE_LOBBY, {
        roomId: null,
        roomRole: null,
        roomPlayerCount: 0,
        hasClientInRoom: false,
        roomStatus: null
      }).catch(() => {});
    }
  }

  /**
   * สร้างห้อง Private Room เมื่อผู้รับกด "ใช่" ตอบรับคำเชิญ
   * - หากผู้ชวนเปิดห้อง Public รออยู่ (และยังไม่มี Client) จะปิดห้อง Public เดิมอย่างปลอดภัยก่อนสร้างห้อง Private
   */
  async hostCreatePrivateRoomForAcceptedInvite(acceptedInv) {
    if (!acceptedInv || this.creatingPrivateRoomForInviteId === acceptedInv.invitationId) return;
    this.creatingPrivateRoomForInviteId = acceptedInv.invitationId;

    const user = this.getCurrentUser();
    if (!user) {
      this.creatingPrivateRoomForInviteId = null;
      return;
    }

    try {
      this.announcePolite(`${acceptedInv.toUsername} ตอบรับคำเชิญแล้ว กำลังสร้างห้องแข่งขันส่วนตัว`);

      // หากผู้ชวนมีห้อง Public ที่เปิดรออยู่เดิม ให้ตรวจสอบว่ายังว่างและปิดอย่างปลอดภัยก่อน
      if (this.currentRoomId) {
        const closedOk = await this.safelyCloseWaitingPublicRoomBeforePrivateTransition();
        if (!closedOk) {
          await this.cancelOutgoingInvitation('cancelled', true);
          this.announcePolite("ไม่สามารถสร้างห้องส่วนตัวได้ เนื่องจากห้องเดิมของคุณมีผู้เล่นเข้าร่วมแล้ว");
          return;
        }
      }

      // ตรวจสอบอีกครั้งว่าเพื่อนยังออนไลน์และพร้อมเข้าห้อง
      const inviteePresence = await getPlayerPresence(acceptedInv.toUserKey);
      if (!inviteePresence.online || inviteePresence.isExpired) {
        await this.cancelOutgoingInvitation('expired', true);
        this.announcePolite(`ผู้เล่น ${acceptedInv.toUsername} หลุดการเชื่อมต่อก่อนเข้าห้องส่วนตัว`);
        return;
      }

      const roomId = await this.generateNextRoomId();
      const myKey = encodeUserKey(user.username);
      const matchId = `${roomId}_${Date.now()}`;
      const now = this.getServerNow();

      const privateRoomData = {
        roomId: roomId,
        matchId: matchId,
        visibility: 'private',
        isPrivate: true,
        invitationId: acceptedInv.invitationId,
        invitedClientKey: acceptedInv.toUserKey,
        invitedClientUsername: acceptedInv.toUsername,
        allowedPlayers: {
          [myKey]: true,
          [acceptedInv.toUserKey]: true
        },
        status: 'waiting',
        createdAt: now,
        updatedAt: now,
        playerCount: 1,
        host: {
          username: user.username,
          userKey: myKey,
          peerId: '',
          connected: true,
          sessionId: this.sessionId,
          heartbeatAt: serverTimestamp(),
          joinedAt: now
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
      await set(roomRef, privateRoomData);

      this.currentRoomId = roomId;
      this.currentRoomVisibility = 'private';
      this.role = 'host';
      this.matchId = matchId;
      this.latestRoomData = privateRoomData;
      this.lastJoinedClientKey = null;
      this.readyAnnouncedForRoomId = null;
      sessionStorage.setItem('pingpong_active_online_room', roomId);

      // ปิดหน้ารายชื่อเพื่อนโดยซิงก์กับ currentRoomId ที่ตั้งค่าแล้ว เพื่อไม่ให้ closeInviteFriendsScreen ยิง cancelOutgoingInvitation ซ้อน
      if (this.invitationBridge && typeof this.invitationBridge.closeInviteFriendsScreen === 'function') {
        if (this.ui.inviteFriendsScreen && !this.ui.inviteFriendsScreen.classList.contains('hidden')) {
          this.invitationBridge.closeInviteFriendsScreen();
        }
      }

      // เล่นเสียงสร้างห้องตามระบบเดิม
      this.sound.play('create', 0, 1.0, false);

      // เปิดหน้า Lobby แสดงข้อมูลห้อง
      this.ensureLobbyVisibleForRoomCard();
      await this.initPeerAndBindRoom(roomId, 'host');

      // อัปเดตโหนดคำเชิญด้วย Transaction เพื่อส่งรหัสห้อง Private ให้ผู้รับเข้าห้องทันที (รองรับกรณี status เป็น pending หรือ accepted)
      const targetInboxRef = ref(this.db, `${INVITATION_CONFIG.INVITATIONS_PATH}/${acceptedInv.toUserKey}`);
      const txRes = await runTransaction(targetInboxRef, (node) => {
        if (!node || node.invitationId !== acceptedInv.invitationId) return;
        if (node.status !== 'accepted' && node.status !== 'pending') return;
        node.status = 'room_ready';
        node.roomId = roomId;
        node.visibility = 'private';
        node.updatedAt = this.getServerNow();
        node.expiresAt = this.getServerNow() + INVITATION_CONFIG.PRIVATE_ROOM_JOIN_WAIT_MS;
        return node;
      }, { applyLocally: false });

      if (!txRes.committed) {
        await this.leaveCurrentRoom(false);
        this.announcePolite("การยืนยันห้องส่วนตัวถูกยกเลิก เนื่องจากคำเชิญเปลี่ยนแปลงสถานะ");
        return;
      }

      // ตั้งเวลารอให้ผู้รับเข้าห้อง Private และเชื่อม PeerJS หากผู้รับหลุดไปก่อนเข้าห้อง ให้ล้างห้อง Private ทันทีไม่ให้ค้าง
      clearTimeout(this.outgoingInviteTimeoutTimer);
      this.outgoingInviteTimeoutTimer = setTimeout(async () => {
        if (this.currentRoomId === roomId && this.role === 'host' && !this.inMatch) {
          const snap = await get(roomRef);
          if (snap.exists() && (!snap.val().client || snap.val().playerCount < 2)) {
            this.announcePolite(`ผู้เล่น ${acceptedInv.toUsername} ไม่สามารถเข้าร่วมห้องส่วนตัวได้ทันเวลา ระบบปิดห้องส่วนตัวอัตโนมัติ`);
            await this.leaveCurrentRoom(false);
            await this.cleanupInvitationNodeIfMatched(acceptedInv.toUserKey, acceptedInv.invitationId);
          }
        }
      }, INVITATION_CONFIG.PRIVATE_ROOM_JOIN_WAIT_MS);
    } catch (err) {
      console.error("Create private room error:", err);
      this.announcePolite("ไม่สามารถสร้างห้องส่วนตัวได้ กรุณาลองใหม่อีกครั้ง");
      await this.cancelOutgoingInvitation('cancelled', true);
    } finally {
      this.creatingPrivateRoomForInviteId = null;
    }
  }

  /**
   * เมื่อผู้รับกดปุ่ม "ใช่" ในหน้าต่างรับคำเชิญ (เรียกจาก ping.html)
   */
  async onAcceptFriendInvitation(invitation, preVerifyResult = null) {
    return this.acceptFriendInvitation(invitation, preVerifyResult);
  }

  async acceptFriendInvitation(invitation, preVerifyResult = null) {
    if (this.acceptingInvitationBusy) return false;
    const user = this.getCurrentUser();
    if (!user || !user.username || !invitation) return false;

    this.acceptingInvitationBusy = true;
    const myKey = encodeUserKey(user.username);
    const hostUsername = invitation.fromUsername || invitation.hostUsername || '';
    const hostKey = invitation.fromUserKey || (hostUsername ? encodeUserKey(hostUsername) : '');
    const myInboxRef = ref(this.db, `${INVITATION_CONFIG.INVITATIONS_PATH}/${myKey}`);

    try {
      this.ensureServerClockSync();
      await this.waitForServerClock(1500);

      // หากผู้รับมีคำเชิญที่ตัวเองส่งค้างไว้ ให้ยกเลิกก่อน
      if (this.outgoingInvitation) {
        await this.cancelOutgoingInvitation('cancelled', true);
      }

      // 1. ตรวจสอบว่าทั้งสองคนยังออนไลน์และว่างจริง
      const latestCheck = await this.verifyCanAcceptInvitationFn({
        inviteeUsername: user.username,
        hostUsername: hostUsername,
        roomId: invitation.roomId || null
      });

      if (!latestCheck || !latestCheck.allowed) {
        this.announcePolite((latestCheck && latestCheck.reason) || "ไม่สามารถตอบรับคำเชิญได้ เนื่องจากผู้เชิญไม่ว่างแล้ว");
        await this.cleanupInvitationNodeIfMatched(myKey, invitation.invitationId);
        return false;
      }

      // 2. งานที่ 4: หากผู้รับกำลังเปิดห้อง Public รออยู่ (เป็น Host ที่ยังไม่มี Client)
      // ต้องตรวจสอบสถานะห้องล่าสุดด้วย Transaction และปิดห้องเดิมอย่างปลอดภัยก่อนย้ายไป Private Room
      if (this.currentRoomId) {
        const closedOldRoom = await this.safelyCloseWaitingPublicRoomBeforePrivateTransition();
        if (!closedOldRoom) {
          this.announcePolite("ไม่สามารถตอบรับคำเชิญได้ เนื่องจากห้องของคุณมีผู้เล่นเข้าร่วมแล้วหรือเริ่มการแข่งขันแล้ว");
          await runTransaction(myInboxRef, (node) => {
            if (!node || node.invitationId !== invitation.invitationId) return;
            node.status = 'busy';
            node.updatedAt = this.getServerNow();
            return node;
          }, { applyLocally: false }).catch(() => {});
          return false;
        }
      }

      // 3. ยืนยันการรับคำเชิญด้วย Firebase Transaction ป้องกันการรับชนกันหรือคำเชิญหมดอายุ
      let txRejectReason = "";
      const acceptTx = await runTransaction(myInboxRef, (node) => {
        const now = this.getServerNow();
        if (!node || typeof node !== 'object') {
          txRejectReason = "ไม่พบคำเชิญนี้ หรือคำเชิญถูกยกเลิกไปแล้ว";
          return;
        }
        if (invitation.invitationId && node.invitationId !== invitation.invitationId) {
          txRejectReason = "คำเชิญนี้ถูกแทนที่หรือหมดอายุไปแล้ว";
          return;
        }
        if (node.status !== 'pending' && node.status !== 'room_ready' && node.status !== 'accepted') {
          txRejectReason = "คำเชิญนี้หมดอายุหรือถูกยกเลิกไปแล้ว";
          return;
        }
        if (typeof node.expiresAt === 'number' && now > node.expiresAt) {
          node.status = 'expired';
          node.updatedAt = now;
          txRejectReason = "คำเชิญหมดอายุแล้ว";
          return node;
        }
        if (node.status === 'pending') {
          node.status = 'accepted';
          node.acceptedAt = now;
          node.updatedAt = now;
          node.expiresAt = now + INVITATION_CONFIG.PRIVATE_ROOM_JOIN_WAIT_MS;
        }
        return node;
      }, { applyLocally: false });

      if (!acceptTx.committed || !acceptTx.snapshot.exists()) {
        this.announcePolite(txRejectReason || "ไม่สามารถตอบรับคำเชิญได้ คำเชิญอาจหมดอายุแล้ว");
        return false;
      }

      const currentInvState = acceptTx.snapshot.val();
      if (currentInvState.status === 'expired') {
        this.announcePolite("คำเชิญหมดอายุแล้ว");
        await this.cleanupInvitationNodeIfMatched(myKey, currentInvState.invitationId);
        return false;
      }

      this.announcePolite(`ตอบรับคำเชิญของ ${hostUsername} แล้ว กำลังเข้าห้องแข่งขันส่วนตัว`);

      // 4. รอรับรหัสห้อง Private Room จากผู้ชวน (กรณีผู้ชวนกำลังสร้างห้อง)
      const privateRoomId = currentInvState.roomId
        ? currentInvState.roomId
        : await this.waitForPrivateRoomReady(myKey, currentInvState.invitationId, hostKey);

      if (!privateRoomId) {
        this.announcePolite(`ไม่สามารถเข้าห้องส่วนตัวของ ${hostUsername} ได้ เนื่องจากผู้เชิญหลุดการเชื่อมต่อหรือยกเลิกห้อง`);
        await this.cleanupInvitationNodeIfMatched(myKey, currentInvState.invitationId);
        return false;
      }

      // 5. เข้าร่วม Private Room ในฐานะ Client พร้อมตรวจสอบสิทธิ์ตัวตนผู้เล่น
      this.ensureLobbyVisibleForRoomCard();
      const joined = await this.joinPrivateRoomAsInvitedClient(privateRoomId, currentInvState.invitationId, hostKey);
      if (!joined) {
        await this.cleanupInvitationNodeIfMatched(myKey, currentInvState.invitationId);
        return false;
      }

      // แจ้งโหนดคำเชิญว่าเข้าห้องสำเร็จแล้ว
      await runTransaction(myInboxRef, (node) => {
        if (!node || node.invitationId !== currentInvState.invitationId) return;
        node.status = 'joined';
        node.updatedAt = this.getServerNow();
        return node;
      }, { applyLocally: false }).catch(() => {});

      this.activeIncomingInvitation = null;
      return true;
    } catch (err) {
      console.error("Accept friend invitation error:", err);
      this.announcePolite("เกิดข้อผิดพลาดในการเข้าห้องส่วนตัว");
      return false;
    } finally {
      this.acceptingInvitationBusy = false;
    }
  }

  /**
   * รอให้ Host สร้างห้อง Private และส่ง roomId กลับมาที่โหนดคำเชิญ
   */
  waitForPrivateRoomReady(myKey, invitationId, hostKey) {
    return new Promise((resolve) => {
      let settled = false;
      let unsubInv = null;
      let timer = null;

      const finish = (roomId) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        if (unsubInv) {
          try { unsubInv(); } catch (e) {}
          unsubInv = null;
        }
        resolve(roomId || null);
      };

      timer = setTimeout(() => {
        finish(null);
      }, INVITATION_CONFIG.PRIVATE_ROOM_JOIN_WAIT_MS);

      const myInboxRef = ref(this.db, `${INVITATION_CONFIG.INVITATIONS_PATH}/${myKey}`);
      unsubInv = onValue(myInboxRef, (snap) => {
        if (!snap.exists()) {
          finish(null);
          return;
        }
        const val = snap.val();
        if (!val || val.invitationId !== invitationId) {
          finish(null);
          return;
        }
        if (val.status === 'cancelled' || val.status === 'expired') {
          finish(null);
          return;
        }
        if ((val.status === 'room_ready' || val.status === 'accepted') && val.roomId) {
          finish(val.roomId);
        }
      }, () => {
        finish(null);
      });
    });
  }

  /**
   * งานที่ 4: ปิดห้อง Public ที่เปิดรออยู่เดิมอย่างปลอดภัยด้วย Transaction
   * ห้ามลบห้องเด็ดขาดหากมีผู้เล่นอื่นเพิ่งเข้าร่วมในจังหวะเดียวกันหรือเริ่มแข่งขันแล้ว
   */
  async safelyCloseWaitingPublicRoomBeforePrivateTransition() {
    if (!this.currentRoomId) return true;
    if (this.inMatch || this.role !== 'host') return false;

    const oldRoomId = this.currentRoomId;
    const roomRef = ref(this.db, `${GAME_ID}/onlineRooms/${oldRoomId}`);

    try {
      let canClose = true;
      const txRes = await runTransaction(roomRef, (room) => {
        if (room === null) return room;
        // หากมี Client เข้าร่วมแล้ว หรือสถานะไม่ใช่ waiting ห้ามลบห้องเด็ดขาด!
        if (room.status !== 'waiting' || (room.client && room.client.userKey) || (room.playerCount && room.playerCount >= 2)) {
          canClose = false;
          return;
        }
        return null;
      }, { applyLocally: false });

      if (!canClose || (!txRes.committed && txRes.snapshot.exists())) {
        return false;
      }

      // ล้าง Listener, Heartbeat และ PeerJS ของห้อง Public เดิมทั้งหมด
      if (this.roomUnsubscribe) {
        try { this.roomUnsubscribe(); } catch (e) {}
        this.roomUnsubscribe = null;
      }
      if (this.presenceDisconnectRef) {
        try { await this.presenceDisconnectRef.cancel(); } catch (e) {}
        this.presenceDisconnectRef = null;
      }
      this.cleanupPeerOnly();
      this.currentRoomId = null;
      this.currentRoomVisibility = 'public';
      this.role = null;
      this.matchId = null;
      this.latestRoomData = null;
      this.lastJoinedClientKey = null;
      this.readyAnnouncedForRoomId = null;
      sessionStorage.removeItem('pingpong_active_online_room');
      this.updateRoomPanelVisibility(false);
      return true;
    } catch (e) {
      console.warn("Close waiting public room before private transition failed:", e);
      return false;
    }
  }

  /**
   * เข้าร่วมห้อง Private Room โดยตรวจสอบสิทธิ์ตัวตนจาก invitedClientKey และ allowedPlayers
   */
  async joinPrivateRoomAsInvitedClient(roomId, invitationId, expectedHostKey = null) {
    const user = this.getCurrentUser();
    if (!user || !roomId) return false;

    const myKey = encodeUserKey(user.username);
    const roomRef = ref(this.db, `${GAME_ID}/onlineRooms/${roomId}`);
    let rejectReason = "";

    try {
      const txResult = await runTransaction(roomRef, (room) => {
        if (room === null) return room;

        const isPrivateRoom = room.visibility === 'private' || room.isPrivate === true;
        if (!isPrivateRoom) {
          rejectReason = "ห้องนี้ไม่ใช่ห้องแข่งขันส่วนตัว";
          return;
        }
        if (room.invitedClientKey && room.invitedClientKey !== myKey) {
          rejectReason = "คุณไม่มีสิทธิ์เข้าร่วมห้องส่วนตัวนี้";
          return;
        }
        if (room.allowedPlayers && !room.allowedPlayers[myKey]) {
          rejectReason = "คุณไม่ได้รับอนุญาตให้เข้าห้องส่วนตัวนี้";
          return;
        }
        if (expectedHostKey && room.host && room.host.userKey && room.host.userKey !== expectedHostKey) {
          rejectReason = "ข้อมูลเจ้าของห้องส่วนตัวไม่ตรงกับผู้เชิญ";
          return;
        }
        if (!room.host || room.host.connected === false) {
          rejectReason = "ผู้เชิญหลุดการเชื่อมต่อจากห้องส่วนตัวแล้ว";
          return;
        }
        if (room.status !== 'waiting' || (room.client && room.client.userKey && room.client.userKey !== myKey)) {
          rejectReason = "ห้องส่วนตัวนี้ไม่ว่างหรือถูกปิดไปแล้ว";
          return;
        }

        const now = this.getServerNow();
        room.client = {
          username: user.username,
          userKey: myKey,
          peerId: '',
          connected: true,
          sessionId: this.sessionId,
          heartbeatAt: now,
          joinedAt: now
        };
        room.playerCount = 2;
        room.status = 'ready';
        room.updatedAt = now;
        return room;
      }, { applyLocally: false });

      if (!txResult.committed || !txResult.snapshot.exists()) {
        this.announcePolite(rejectReason || "ไม่สามารถเข้าร่วมห้องส่วนตัวได้");
        return false;
      }

      const updatedRoom = txResult.snapshot.val();
      if (!updatedRoom.client || updatedRoom.client.userKey !== myKey) {
        this.announcePolite("ไม่สามารถเข้าร่วมห้องส่วนตัวได้");
        return false;
      }

      this.currentRoomId = roomId;
      this.currentRoomVisibility = 'private';
      this.role = 'client';
      this.matchId = updatedRoom.matchId;
      this.latestRoomData = updatedRoom;
      this.lastJoinedClientKey = myKey;
      this.readyAnnouncedForRoomId = null;
      sessionStorage.setItem('pingpong_active_online_room', roomId);

      // เล่นเสียงเข้าร่วมห้องสำเร็จตามระบบเดิม
      this.sound.play('select', 0, 1.0, false);

      await this.initPeerAndBindRoom(roomId, 'client');
      this.announcePolite(`เข้าห้องส่วนตัว ${roomId} สำเร็จ พบกับ ${updatedRoom.host.username} กำลังเชื่อมต่อสัญญาณ`);
      return true;
    } catch (err) {
      console.error("Join private room error:", err);
      this.announcePolite("ไม่สามารถเชื่อมต่อเข้าห้องส่วนตัวได้");
      return false;
    }
  }

  /**
   * งานที่ 5: เมื่อผู้รับกดปุ่ม "ไม่" ปฏิเสธคำเชิญ (เรียกจาก ping.html)
   */
  async onDeclineFriendInvitation(invitation) {
    return this.declineFriendInvitation(invitation);
  }

  async declineFriendInvitation(invitation) {
    const user = this.getCurrentUser();
    this.activeIncomingInvitation = null;
    if (!user || !user.username) return false;

    const myKey = encodeUserKey(user.username);
    const myInboxRef = ref(this.db, `${INVITATION_CONFIG.INVITATIONS_PATH}/${myKey}`);

    try {
      await runTransaction(myInboxRef, (node) => {
        if (!node || typeof node !== 'object') return node;
        if (invitation && invitation.invitationId && node.invitationId !== invitation.invitationId) {
          return;
        }
        node.status = 'declined';
        node.updatedAt = this.getServerNow();
        return node;
      }, { applyLocally: false });
      return true;
    } catch (e) {
      return false;
    }
  }

  ensureLobbyVisibleForRoomCard() {
    this.activeMode = true;
    this.inMatch = false;

    if (this.ui.homeScreen) this.ui.homeScreen.classList.add('hidden');
    if (this.ui.rulesScreen) this.ui.rulesScreen.classList.add('hidden');
    if (this.ui.roundSelectScreen) this.ui.roundSelectScreen.classList.add('hidden');
    if (this.ui.leaderboardScreen) this.ui.leaderboardScreen.classList.add('hidden');
    if (this.ui.gameScreen) this.ui.gameScreen.classList.add('hidden');
    if (this.ui.inviteFriendsScreen) this.ui.inviteFriendsScreen.classList.add('hidden');
    if (this.ui.backFromInviteFriendsBtn) this.ui.backFromInviteFriendsBtn.classList.add('hidden');

    if (this.ui.startBtn) this.ui.startBtn.classList.add('hidden');
    if (this.ui.onlineModeBtn) this.ui.onlineModeBtn.classList.add('hidden');
    if (this.ui.selectRoundBtn) this.ui.selectRoundBtn.classList.add('hidden');
    if (this.ui.rulesBtn) this.ui.rulesBtn.classList.add('hidden');
    if (this.ui.inGameLeaderboardBtn) this.ui.inGameLeaderboardBtn.classList.add('hidden');

    if (this.ui.btnHomeMain) {
      this.ui.btnHomeMain.classList.remove('hidden');
    }
    if (this.ui.onlineLobbyScreen) {
      this.ui.onlineLobbyScreen.classList.remove('hidden');
    }
    this.subscribeWaitingRooms();
    this.loadPlayerOnlineStatsUI().catch(() => {});
  }

  async openLobby() {
    this.activeMode = true;
    this.inMatch = false;
    this.ui.homeScreen.classList.add('hidden');
    this.ui.rulesScreen.classList.add('hidden');
    this.ui.roundSelectScreen.classList.add('hidden');
    this.ui.leaderboardScreen.classList.add('hidden');
    this.ui.gameScreen.classList.add('hidden');
    if (this.ui.inviteFriendsScreen) {
      this.ui.inviteFriendsScreen.classList.add('hidden');
    }
    if (this.ui.backFromInviteFriendsBtn) {
      this.ui.backFromInviteFriendsBtn.classList.add('hidden');
    }

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
    if (this.outgoingInvitation) {
      await this.cancelOutgoingInvitation('cancelled', true);
    }
    await this.leaveCurrentRoom(true);
    this.unsubscribeWaitingRooms();
    this.activeMode = false;
    this.inMatch = false;

    this.ui.onlineLobbyScreen.classList.add('hidden');
    if (this.ui.inviteFriendsScreen) {
      this.ui.inviteFriendsScreen.classList.add('hidden');
    }
    if (this.ui.backFromInviteFriendsBtn) {
      this.ui.backFromInviteFriendsBtn.classList.add('hidden');
    }
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
   * รับข้อมูลรายการห้องล่าสุด: แสดงเฉพาะห้อง Public ที่ยังใช้งานได้ และสั่งลบห้องที่ยืนยันว่าหมดอายุออกจาก Firebase
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

  isPrivateRoomRecord(room) {
    if (!room || typeof room !== 'object') return false;
    return room.visibility === 'private' || room.isPrivate === true;
  }

  /**
   * ประเมินห้อง waiting ด้วยเวลา Firebase Server
   * joinable = แสดงปุ่มเข้าร่วมในรายการห้องสาธารณะได้ (ต้องไม่ใช่ห้อง private)
   * expired = ยืนยันว่าหมดอายุและลบได้
   */
  evaluateWaitingRoom(room, now) {
    const C = ROOM_PRESENCE_CONFIG;
    const num = (v) => (typeof v === 'number' && isFinite(v) ? v : null);
    if (!room || typeof room !== 'object') return { joinable: false, expired: false };
    if (room.status !== 'waiting') return { joinable: false, expired: false };

    const host = room.host;
    if (!host || typeof host !== 'object') return { joinable: false, expired: true };
    if (room.client) return { joinable: false, expired: false };

    const isPrivate = this.isPrivateRoomRecord(room);
    const hb = num(host.heartbeatAt);
    const dc = num(host.disconnectedAt);

    // ห้องที่ใช้ระบบ Presence ใหม่ (เวลาจาก Firebase Server)
    if (hb !== null) {
      const evidence = Math.max(hb, dc !== null ? dc : 0);
      const age = now - evidence;
      if (host.connected === false) {
        return { joinable: false, expired: age > C.DISCONNECT_GRACE_MS };
      }
      const isExpired = age > C.HEARTBEAT_STALE_MS;
      return {
        joinable: !isPrivate && !isExpired,
        expired: isExpired
      };
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
    const isLegacyExpired = age > C.LEGACY_STALE_MS;
    return {
      joinable: !isPrivate && !isLegacyExpired,
      expired: isLegacyExpired
    };
  }

  /**
   * คืนรายการเฉพาะห้อง Public ที่เปิดรอผู้เล่น (กรองห้อง Private ออก 100%)
   */
  getJoinableRooms(roomsObj) {
    const now = this.getServerNow();
    return Object.values(roomsObj || {}).filter((r) => {
      if (!r || typeof r !== 'object' || !r.roomId) return false;
      if (this.isPrivateRoomRecord(r)) return false;
      return r.status === 'waiting' && r.host && !r.client &&
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
        this.currentRoomVisibility = this.isPrivateRoomRecord(room) ? 'private' : 'public';
        this.role = 'host';
        this.matchId = room.matchId;
        this.latestRoomData = room;
        this.lastJoinedClientKey = room.client ? room.client.userKey : null;
        this.readyAnnouncedForRoomId = null;
        await this.initPeerAndBindRoom(savedRoomId, 'host');
        this.announcePolite(`กู้คืนห้อง ${savedRoomId} สำเร็จในฐานะหัวหน้าห้อง`, true);
      } else if (room.client && room.client.userKey === myKey) {
        this.currentRoomId = savedRoomId;
        this.currentRoomVisibility = this.isPrivateRoomRecord(room) ? 'private' : 'public';
        this.role = 'client';
        this.matchId = room.matchId;
        this.latestRoomData = room;
        this.lastJoinedClientKey = room.client.userKey;
        this.readyAnnouncedForRoomId = null;
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
        visibility: 'public',
        isPrivate: false,
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
      this.currentRoomVisibility = 'public';
      this.role = 'host';
      this.matchId = matchId;
      this.latestRoomData = roomData;
      this.lastJoinedClientKey = null;
      this.readyAnnouncedForRoomId = null;
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
   * เข้าร่วมห้องสาธารณะโดยใช้ Firebase Transaction ป้องกันผู้เล่นคนที่ 3 หรือการแอบเข้าห้อง Private โดยไม่ได้รับเชิญ
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
        // ป้องกันการเข้าห้อง Private ผ่านฟังก์ชันห้องสาธารณะ यदिไม่ได้รับอนุญาตใน allowedPlayers / invitedClientKey
        if (this.isPrivateRoomRecord(room)) {
          const isAuthorizedClient = (room.invitedClientKey && room.invitedClientKey === myKey) ||
            (room.allowedPlayers && room.allowedPlayers[myKey] === true);
          if (!isAuthorizedClient) {
            rejectReason = "ห้องนี้เป็นห้องส่วนตัว เฉพาะผู้เล่นที่ได้รับคำเชิญเท่านั้นที่เข้าร่วมได้";
            return;
          }
        }
        if (
          this.serverClockReady &&
          room.status === 'waiting' &&
          !room.client &&
          !this.isPrivateRoomRecord(room) &&
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
      this.currentRoomVisibility = this.isPrivateRoomRecord(updatedRoom) ? 'private' : 'public';
      this.role = 'client';
      this.matchId = updatedRoom.matchId;
      this.latestRoomData = updatedRoom;
      this.lastJoinedClientKey = myKey;
      this.readyAnnouncedForRoomId = null;
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

    // ผูก Firebase Room Listener ทันทีเพื่อรับข้อมูลสถานะห้องแบบ Real-time
    this.subscribeRoomChanges(roomId);

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
      if (this.currentRoomId !== roomId || this.role !== role) return;
      this.peerReady = true;
      this.localPeerId = peerId;

      // ลงทะเบียน onDisconnect ก่อน แล้วจึงประกาศสถานะออนไลน์ + เริ่ม Heartbeat
      this.startPresence(roomId, role);

      // แก้ปัญหา Race Condition: หาก Client ได้รับข้อมูลห้องจาก Firebase ก่อนที่ PeerJS ของ Client จะพร้อม
      // ให้เริ่มเชื่อมต่อไปหา Host ทันทีที่ PeerJS เปิดสำเร็จ
      if (
        this.role === 'client' &&
        !this.conn &&
        this.latestRoomData &&
        this.latestRoomData.host &&
        this.latestRoomData.host.peerId &&
        this.latestRoomData.host.connected !== false
      ) {
        this.connectToHostPeer(this.latestRoomData.host.peerId);
      }
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

        // 2) ประกาศสถานะออนไลน์และบันทึก peerId ลง Firebase อย่างแน่นอน
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
      const now = this.getServerNow();
      const patch = {
        sessionId: this.sessionId,
        connected: true,
        heartbeatAt: now,
        lastSeen: now,
        disconnectedAt: null
      };
      if (this.localPeerId) {
        patch.peerId = this.localPeerId;
      }

      // ใช้ update ร่วมกับตรวจสอบการมีอยู่ของโหนด เพื่อหลีกเลี่ยงปัญหา runTransaction ชนกับ set/runTransaction ของห้องเมื่อเพิ่งสร้างหรือเข้าร่วมห้องใหม่
      const snap = await get(roleRef);
      if (!snap.exists()) return 'lost';

      await update(roleRef, patch);
      return 'ok';
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
        if (this.localPeerId && !node.peerId) {
          node.peerId = this.localPeerId;
        }
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
    this.latestRoomData = room;
    this.hostUsername = room.host ? room.host.username : '-';
    if (room.client && room.client.username) {
      this.clientUsername = room.client.username;
    } else if (!this.clientUsername) {
      this.clientUsername = 'รอผู้เล่นเข้าร่วม...';
    }
    this.currentRoomVisibility = this.isPrivateRoomRecord(room) ? 'private' : 'public';

    const hasBothPlayers = this.hasBothPlayersInRoom(room);

    // หยุด Timer รอเพื่อนเข้าห้อง Private ทันทีที่เห็นว่า Client เข้าห้องใน Firebase แล้ว
    if (hasBothPlayers && this.outgoingInviteTimeoutTimer) {
      clearTimeout(this.outgoingInviteTimeoutTimer);
      this.outgoingInviteTimeoutTimer = null;
    }

    // หากตัวเราเปิด PeerJS สำเร็จแล้ว แต่ใน Snapshot ของห้องยังไม่มี peerId ของเรา (เช่น ถูกเขียนทับตอนสร้าง/อัปเดตห้อง) ให้ซิงก์กลับทันที
    const myNode = this.role === 'host' ? room.host : room.client;
    if (this.peerReady && this.localPeerId && myNode && myNode.peerId !== this.localPeerId) {
      update(ref(this.db, `${GAME_ID}/onlineRooms/${room.roomId}/${this.role}`), {
        peerId: this.localPeerId,
        connected: true,
        updatedAt: this.getServerNow()
      }).catch(() => {});
    }

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
    const roomTypeSuffix = this.currentRoomVisibility === 'private' ? ' (ห้องส่วนตัว)' : '';
    if (this.ui.roomCodeDisplay) {
      this.ui.roomCodeDisplay.textContent = `รหัสห้อง: ${room.roomId}${roomTypeSuffix}`;
    }
    if (this.ui.roomHostDisplay) {
      this.ui.roomHostDisplay.textContent = `หัวหน้าห้อง (Host): ${this.hostUsername}`;
    }
    if (this.ui.roomClientDisplay) {
      this.ui.roomClientDisplay.textContent = `ผู้ท้าชิง (Client): ${this.clientUsername}`;
    }

    // ซิงก์สถานะ Presence ของผู้เล่นกับ auth.js ให้ตรงกับสถานะห้องจริงเสมอ
    if (!this.inMatch) {
      if (room.status === 'playing') {
        this.updatePresenceStatusFn(this.presenceStates.PLAYING_ONLINE, {
          roomId: room.roomId,
          roomRole: this.role,
          roomPlayerCount: 2,
          hasClientInRoom: true,
          roomStatus: 'playing'
        }).catch(() => {});
      } else if (hasBothPlayers || room.status === 'ready') {
        this.updatePresenceStatusFn(this.presenceStates.ONLINE_READY, {
          roomId: room.roomId,
          roomRole: this.role,
          roomPlayerCount: 2,
          hasClientInRoom: true,
          roomStatus: 'ready'
        }).catch(() => {});
      } else if (room.status === 'waiting' && this.role === 'host') {
        // ห้อง Private ที่กำลังรอเพื่อนเข้า จะไม่เปิดรับคำเชิญซ้อนจากคนอื่น
        const waitingState = this.currentRoomVisibility === 'private'
          ? this.presenceStates.INVITATION_PENDING
          : this.presenceStates.WAITING_ROOM;
        this.updatePresenceStatusFn(waitingState, {
          roomId: room.roomId,
          roomRole: 'host',
          roomPlayerCount: 1,
          hasClientInRoom: false,
          roomStatus: 'waiting'
        }).catch(() => {});
      }
    }

    // หากผู้เล่นเป็น Client และเห็น Peer ID ของ Host แล้ว แต่ยังไม่ได้เชื่อม DataChannel (หรือ Host เปลี่ยน Peer ID) ให้เริ่มเชื่อมทันที
    if (
      this.role === 'client' &&
      this.peerReady &&
      room.host &&
      room.host.peerId &&
      room.host.connected !== false
    ) {
      const needConnect = !this.conn || (!this.channelReady && this.connectingHostPeerId !== room.host.peerId);
      if (needConnect) {
        this.connectToHostPeer(room.host.peerId);
      }
    }

    // กรณี Fallback สำหรับ Client: หาก Host กดเริ่มเกมและอัปเดตสถานะห้องใน Firebase เป็น playing แล้ว แต่ข้อความ match_start ผ่าน PeerJS ตกหล่น
    if (
      this.role === 'client' &&
      !this.inMatch &&
      room.status === 'playing' &&
      this.channelReady &&
      room.matchState
    ) {
      this.matchId = room.matchId || this.matchId;
      this.hostSetsWon = room.matchState.hostSetsWon || 0;
      this.clientSetsWon = room.matchState.clientSetsWon || 0;
      this.pointEventSeq = 0;
      this.lastProcessedPointEventId = null;
      this.lastPlayedMissAudioEventId = null;
      this.enterMatchScreenUI();
      this.startOnlineSet(room.matchState.currentSetNumber || 1, room.matchState.initialServer || 'host');
      return;
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
    } else if (!this.inMatch && hasBothPlayers && oppNode && oppNode.connected === false) {
      // กรณีอยู่ในห้องรอ (รวมถึงห้อง Private) แล้วอีกฝ่ายปิดหน้าเว็บหรือหลุดการเชื่อมต่อก่อนเริ่มเกม
      if (this.currentRoomVisibility === 'private') {
        const oppLabel = this.role === 'host' ? this.clientUsername : this.hostUsername;
        this.announcePolite(`ผู้เล่น ${oppLabel} ออกจากห้องหรือหลุดการเชื่อมต่อ ระบบกำลังปิดห้องส่วนตัว`, true);
        this.leaveCurrentRoom(false);
        return;
      }
    }

    // อัปเดตสถานะในหน้าห้องรอและปุ่มเริ่มการแข่งขันของ Host ให้ตรงกับสถานะจริง
    if (!this.inMatch) {
      this.syncRoomReadyUIAndAnnounce(room);
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
    if (this.channelReady && this.conn && this.conn.open && this.connectingHostPeerId === hostPeerId) {
      return;
    }

    clearTimeout(this.peerConnectTimer);
    this.connectingHostPeerId = hostPeerId;

    const outgoingConn = this.peer.connect(hostPeerId, {
      reliable: true,
      serialization: 'json'
    });

    this.setupDataConnection(outgoingConn);

    this.peerConnectTimer = setTimeout(() => {
      if (!this.channelReady) {
        this.connectingHostPeerId = null;
        this.announcePolite(
          "ไม่สามารถเชื่อมต่อ PeerJS โดยตรงได้ภายในเวลาที่กำหนด เครือข่ายอาจมีไฟร์วอลล์หรือต้องตั้งค่า TURN Server กรุณาลองสร้างห้องใหม่หรือสลับเครือข่าย",
          true
        );
        if (this.currentRoomVisibility === 'private' && !this.inMatch) {
          this.leaveCurrentRoom(false);
        } else {
          this.syncRoomReadyUIAndAnnounce();
        }
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

    const onChannelOpen = () => {
      if (this.conn !== connection) return;
      clearTimeout(this.peerConnectTimer);
      clearTimeout(this.outgoingInviteTimeoutTimer);
      this.channelReady = true;
      this.startPingLoop();

      // ส่งข้อมูล Handshake ทันทีที่ DataChannel เปิด เพื่อให้ทั้งสองฝั่งรับทราบชื่อผู้เล่นและสถานะความพร้อมทันทีแม้ Firebase Snapshot จะมาช้ากว่าเล็กน้อย
      const user = this.getCurrentUser();
      this.sendPeerMessage({
        type: 'peer_ready_handshake',
        roomId: this.currentRoomId,
        role: this.role,
        username: user && user.username ? user.username : (this.role === 'host' ? this.hostUsername : this.clientUsername),
        userKey: user && user.username ? encodeUserKey(user.username) : null
      });

      if (this.opponentDisconnected) {
        this.handleOpponentReconnected();
      } else {
        this.syncRoomReadyUIAndAnnounce();
      }
    };

    connection.on('open', onChannelOpen);
    if (connection.open) {
      onChannelOpen();
    }

    connection.on('data', (data) => {
      this.handlePeerData(data);
    });

    connection.on('close', () => {
      if (this.conn !== connection) return;
      this.channelReady = false;
      this.connectingHostPeerId = null;
      if (this.inMatch && this.matchStatus !== 'finished' && this.matchStatus !== 'aborted') {
        this.handleOpponentDisconnect();
      } else if (!this.inMatch) {
        this.syncRoomReadyUIAndAnnounce();
      }
    });

    connection.on('error', (err) => {
      console.warn("DataChannel error:", err);
      if (!this.inMatch) {
        this.syncRoomReadyUIAndAnnounce();
      }
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
    if (this.startingMatchBusy || this.inMatch) return;

    this.startingMatchBusy = true;
    if (this.ui.startOnlineMatchBtn) {
      this.ui.startOnlineMatchBtn.disabled = true;
    }

    try {
      const roomRef = ref(this.db, `${GAME_ID}/onlineRooms/${this.currentRoomId}`);
      const snap = await get(roomRef);
      if (!snap.exists()) {
        this.announcePolite("ไม่พบข้อมูลห้อง");
        return;
      }

      const room = snap.val();
      this.latestRoomData = room;

      if (!this.hasBothPlayersInRoom(room)) {
        this.announcePolite("ยังไม่สามารถเริ่มได้ ต้องรอผู้เล่นครบ 2 คนก่อน");
        this.syncRoomReadyUIAndAnnounce(room);
        return;
      }

      if (!this.channelReady || !this.conn || !this.conn.open) {
        this.announcePolite("ยังไม่สามารถเริ่มได้ การเชื่อมต่อ PeerJS ยังไม่พร้อม");
        this.syncRoomReadyUIAndAnnounce(room);
        return;
      }

      if (room.status === 'playing' && this.inMatch) {
        return;
      }

      this.hostUsername = room.host ? room.host.username : this.hostUsername;
      this.clientUsername = room.client ? room.client.username : this.clientUsername;
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
    } catch (err) {
      console.error("Host start match error:", err);
      this.announcePolite("เกิดข้อผิดพลาดในการเริ่มการแข่งขัน กรุณาลองอีกครั้ง");
      this.syncRoomReadyUIAndAnnounce();
    } finally {
      this.startingMatchBusy = false;
      if (!this.inMatch && this.ui.startOnlineMatchBtn) {
        this.ui.startOnlineMatchBtn.disabled = false;
      }
    }
  }

  enterMatchScreenUI() {
    this.inMatch = true;
    if (this.ui.btnHomeMain) {
      this.ui.btnHomeMain.classList.add('hidden');
    }
    this.ui.onlineLobbyScreen.classList.add('hidden');
    if (this.ui.inviteFriendsScreen) {
      this.ui.inviteFriendsScreen.classList.add('hidden');
    }
    if (this.ui.backFromInviteFriendsBtn) {
      this.ui.backFromInviteFriendsBtn.classList.add('hidden');
    }
    if (this.ui.backFromLobbyBtn) {
      this.ui.backFromLobbyBtn.classList.add('hidden');
    }
    this.ui.startOnlineMatchBtn.classList.add('hidden');

    this.ui.gameScreen.classList.remove('hidden');
    if (this.ui.leaveRoomBtn) {
      this.ui.leaveRoomBtn.classList.add('hidden');
    }

    this.updatePresenceStatusFn(this.presenceStates.PLAYING_ONLINE, {
      roomId: this.currentRoomId || null,
      roomRole: this.role,
      roomPlayerCount: 2,
      hasClientInRoom: true,
      roomStatus: 'playing'
    }).catch(() => {});

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
      case 'peer_ready_handshake': {
        if (this.role === 'host' && data.role === 'client') {
          if (data.username) {
            this.clientUsername = data.username;
            if (this.ui.roomClientDisplay) {
              this.ui.roomClientDisplay.textContent = `ผู้ท้าชิง (Client): ${this.clientUsername}`;
            }
            if (this.latestRoomData) {
              this.latestRoomData.client = {
                username: data.username,
                userKey: data.userKey || encodeUserKey(data.username),
                connected: true
              };
              this.latestRoomData.playerCount = 2;
            }
          }
        } else if (this.role === 'client' && data.role === 'host') {
          if (data.username && (!this.hostUsername || this.hostUsername === '-')) {
            this.hostUsername = data.username;
            if (this.ui.roomHostDisplay) {
              this.ui.roomHostDisplay.textContent = `หัวหน้าห้อง (Host): ${this.hostUsername}`;
            }
          }
        }
        if (!this.inMatch) {
          this.syncRoomReadyUIAndAnnounce();
        }
        break;
      }

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
          if (!this.inMatch && this.channelReady) {
            this.syncRoomReadyUIAndAnnounce();
          }
        }
        break;
      }

      case 'match_start': {
        if (this.role === 'client' && !this.inMatch) {
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

      // หากเป็นห้อง Private ที่จบการแข่งขันสมบูรณ์แล้ว ให้ลบข้อมูลห้องออกจาก Firebase เพื่อไม่ให้ค้างสะสม
      if (this.currentRoomId && this.currentRoomVisibility === 'private' && this.role === 'host') {
        const finishedRoomId = this.currentRoomId;
        setTimeout(() => {
          remove(ref(this.db, `${GAME_ID}/onlineRooms/${finishedRoomId}`)).catch(() => {});
        }, 5000);
      }
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
      if (this.ui.inviteFriendsSection) {
        this.ui.inviteFriendsSection.classList.add('hidden');
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
      if (this.ui.inviteFriendsSection) {
        this.ui.inviteFriendsSection.classList.remove('hidden');
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
    this.connectingHostPeerId = null;
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
    const leavingVisibility = this.currentRoomVisibility;
    const wasInMatch = this.inMatch;
    const prevStatus = this.matchStatus;

    this.currentRoomId = null;
    this.currentRoomVisibility = 'public';
    this.role = null;
    this.latestRoomData = null;
    this.inMatch = false;
    this.startingMatchBusy = false;
    this.opponentDisconnected = false;
    this.canLocalPlayerHit = false;
    this.lastJoinedClientKey = null;
    this.readyAnnouncedForRoomId = null;
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
          const isPrivate = leavingVisibility === 'private' || this.isPrivateRoomRecord(room);
          if (room.status === 'waiting' && leavingRole === 'host') {
            await this.removeRoomIfOwned(roomRef);
          } else if (room.status === 'ready' && !wasInMatch) {
            if (leavingRole === 'host' || isPrivate) {
              // หากเป็นห้อง Private แล้วฝ่ายใดฝ่ายหนึ่งออกจากห้องก่อนเริ่มแข่ง ให้ปิดห้อง Private ทันทีเพื่อไม่ให้ค้าง
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
      if (this.ui.btnHomeMain) {
        this.ui.btnHomeMain.classList.remove('hidden');
      }
      await this.loadPlayerOnlineStatsUI();
    }

    if (this.activeMode && !this.inMatch) {
      this.updatePresenceStatusFn(this.presenceStates.ONLINE_LOBBY, {
        roomId: null,
        roomRole: null,
        roomPlayerCount: 0,
        hasClientInRoom: false,
        roomStatus: null
      }).catch(() => {});
    }

    if (notifyUser) {
      this.announcePolite("ออกจากห้องออนไลน์เรียบร้อยแล้ว", true);
    }
  }
}