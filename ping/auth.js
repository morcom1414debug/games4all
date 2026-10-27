/**
 * auth.js
 * ระบบยืนยันตัวตนและจัดการเซสชันผู้เล่น (PingPong Realtime Database + Session Persistence)
 * รองรับการจำสถานะข้าม Safari iOS / In-App Browser และทนต่อการกด Refresh 100%
 * พร้อมระบบติดตามสถานะผู้เล่นออนไลน์แบบ Real-time (Player Presence)
 */
import { initializeApp, getApps, getApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { 
  getDatabase, 
  ref, 
  get, 
  set, 
  update,
  onValue,
  onDisconnect,
  runTransaction,
  serverTimestamp
} from "https://www.gstatic.com/firebasejs/10.8.0/firebase-database.js";

const firebaseConfig = {
  apiKey: "AIzaSyDvcdgsyT5sDdYTYKIqetzNL9Be-MFC0l4",
  authDomain: "xo-game-134ec.firebaseapp.com",
  databaseURL: "https://xo-game-134ec-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "xo-game-134ec",
  storageBucket: "xo-game-134ec.firebasestorage.app",
  messagingSenderId: "318375224157",
  appId: "1:318375224157:web:35a997751e82f7b6b77eb4"
};

const app = !getApps().length ? initializeApp(firebaseConfig) : getApp();
export const database = getDatabase(app);
export const GAME_ID = "PingPong";

export function encodeUserKey(username) {
  return encodeURIComponent(username.trim().toLowerCase()).replace(/\./g, '%2E');
}

// ระบบจัดการเซสชันแบบไฮบริด (LocalStorage + Cookie สำหรับ Safari iOS)
export const SessionStorage = {
  save(user) {
    try {
      localStorage.setItem("pingpong_session_user", JSON.stringify(user));
    } catch (e) {}
    const d = new Date();
    d.setTime(d.getTime() + (365 * 24 * 60 * 60 * 1000));
    document.cookie = `pingpong_user_key=${encodeUserKey(user.username)}; expires=${d.toUTCString()}; path=/; SameSite=Lax`;
  },
  getSavedKey() {
    try {
      const item = localStorage.getItem("pingpong_session_user");
      if (item) {
        const u = JSON.parse(item);
        if (u && u.username) return encodeUserKey(u.username);
      }
    } catch (e) {}
    const nameEQ = "pingpong_user_key=";
    const ca = document.cookie.split(';');
    for (let c of ca) {
      c = c.trim();
      if (c.indexOf(nameEQ) === 0) return decodeURIComponent(c.substring(nameEQ.length));
    }
    return null;
  },
  clear() {
    try {
      localStorage.removeItem("pingpong_session_user");
    } catch (e) {}
    document.cookie = "pingpong_user_key=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/; SameSite=Lax";
  }
};

/**
 * =========================================================================
 * ระบบติดตามสถานะผู้เล่นแบบ Real-time (Player Presence)
 * โครงสร้างข้อมูลแยกเก็บที่: PingPong/presence/{userKey}
 * =========================================================================
 */
export const PRESENCE_STATES = Object.freeze({
  HOME: "home",
  RULES: "rules",
  ROUND_SELECT: "round_select",
  PLAYING_BOT: "playing_bot",
  ONLINE_LOBBY: "online_lobby",
  WAITING_ROOM: "waiting_room",
  ONLINE_READY: "online_ready",
  PLAYING_ONLINE: "playing_online",
  INVITATION_PENDING: "invitation_pending",
  OFFLINE: "offline"
});

export const INVITABLE_STATES = Object.freeze([
  PRESENCE_STATES.HOME,
  PRESENCE_STATES.RULES,
  PRESENCE_STATES.ROUND_SELECT,
  PRESENCE_STATES.PLAYING_BOT,
  PRESENCE_STATES.ONLINE_LOBBY,
  PRESENCE_STATES.WAITING_ROOM
]);

export const PRESENCE_CONFIG = Object.freeze({
  HEARTBEAT_INTERVAL_MS: 20000,
  HEARTBEAT_STALE_MS: 70000,
  DISCONNECT_GRACE_MS: 25000,
  SWEEP_INTERVAL_MS: 15000
});

const VALID_STATE_SET = new Set(Object.values(PRESENCE_STATES));
const INVITABLE_STATE_SET = new Set(INVITABLE_STATES);

export function evaluatePresenceRecord(record, serverNow = Date.now(), inviterUsernameOrKey = null) {
  if (!record || typeof record !== "object" || !record.username) {
    return { online: false, isExpired: true, isInvitable: false, status: PRESENCE_STATES.OFFLINE, reason: "ไม่พบข้อมูลสถานะผู้เล่น" };
  }

  const username = String(record.username).trim();
  const userKey = record.userKey || encodeUserKey(username);
  const hb = typeof record.heartbeatAt === "number" ? record.heartbeatAt : 0;
  const up = typeof record.updatedAt === "number" ? record.updatedAt : 0;
  const dc = typeof record.disconnectedAt === "number" ? record.disconnectedAt : 0;
  const lastActive = Math.max(hb, up, dc);

  const ageMs = lastActive > 0 ? (serverNow - lastActive) : Infinity;
  const isDisconnected = record.online === false || record.connected === false || record.status === PRESENCE_STATES.OFFLINE;
  const isExpired = isDisconnected
    ? ageMs > PRESENCE_CONFIG.DISCONNECT_GRACE_MS
    : ageMs > PRESENCE_CONFIG.HEARTBEAT_STALE_MS;
  const isOnline = !isDisconnected && !isExpired;

  const status = isOnline && VALID_STATE_SET.has(record.status) ? record.status : PRESENCE_STATES.OFFLINE;
  const inviterKey = inviterUsernameOrKey ? encodeUserKey(String(inviterUsernameOrKey)) : null;
  const isSelf = Boolean(inviterKey && inviterKey === userKey);

  const hasClient = Boolean(record.hasClientInRoom || (record.roomPlayerCount || 0) >= 2);
  let isInvitable = false;
  let reason = "สามารถเชิญได้";

  if (!isOnline) {
    reason = isExpired ? "สถานะผู้เล่นหมดอายุ" : "ผู้เล่นออฟไลน์";
  } else if (isSelf) {
    reason = "ไม่สามารถเชิญตัวเองได้";
  } else if (status === PRESENCE_STATES.PLAYING_ONLINE) {
    reason = "กำลังแข่งขันออนไลน์กับผู้เล่นจริง";
  } else if (status === PRESENCE_STATES.ONLINE_READY || hasClient) {
    reason = "อยู่ในห้องออนไลน์ที่มีคู่แข่งขันแล้ว";
  } else if (status === PRESENCE_STATES.INVITATION_PENDING) {
    reason = "กำลังดำเนินการตอบรับคำเชิญอื่น";
  } else if (status === PRESENCE_STATES.WAITING_ROOM) {
    if ((record.roomRole && record.roomRole !== "host") || (record.roomStatus && record.roomStatus !== "waiting")) {
      reason = "ห้องไม่ว่างหรือเริ่มเกมแล้ว";
    } else {
      isInvitable = true;
    }
  } else if (INVITABLE_STATE_SET.has(status)) {
    isInvitable = true;
  } else {
    reason = "สถานะไม่พร้อมรับคำเชิญ";
  }

  return {
    username,
    userKey,
    status,
    online: isOnline,
    isExpired,
    isInvitable,
    reason,
    roomId: record.roomId || null,
    roomRole: record.roomRole || null,
    roomPlayerCount: record.roomPlayerCount || 0,
    hasClientInRoom: hasClient,
    roomStatus: record.roomStatus || null,
    checkpoint: record.checkpoint || "รอบ 16 คนสุดท้าย",
    wins: record.wins || 0,
    points: record.points || 0,
    sessionId: record.sessionId || null,
    heartbeatAt: hb || null,
    updatedAt: up || null
  };
}

class PlayerPresenceService {
  constructor(db) {
    this.db = db;
    this.sessionId = `ps_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`;
    this.user = null;
    this.userKey = null;
    this.status = PRESENCE_STATES.HOME;
    this.meta = { roomId: null, roomRole: null, roomPlayerCount: 0, hasClientInRoom: false, roomStatus: null };

    this.active = false;
    this.token = 0;
    this.firebaseConnected = false;
    this.serverOffset = 0;
    this.offsetUnsub = null;
    this.connUnsub = null;
    this.onDisconnectRef = null;
    this.heartbeatTimer = null;
    this.wakeHandler = null;

    this.subscribers = new Map();
    this.subSeq = 1;
    this.listUnsub = null;
    this.sweepTimer = null;
    this.rawCache = {};
  }

  ensureClockSync() {
    if (this.offsetUnsub) return;
    try {
      this.offsetUnsub = onValue(ref(this.db, ".info/serverTimeOffset"), (snap) => {
        const val = snap.val();
        if (typeof val === "number" && isFinite(val)) this.serverOffset = val;
      }, () => {});
    } catch (e) {}
  }

  getServerNow() {
    return Date.now() + (this.serverOffset || 0);
  }

  setPublicUser(user) {
    if (!user || !user.username) return false;
    const username = String(user.username).trim();
    this.user = {
      username,
      checkpoint: user.checkpoint || "รอบ 16 คนสุดท้าย",
      roundIndex: typeof user.roundIndex === "number" ? user.roundIndex : 0,
      wins: typeof user.wins === "number" ? user.wins : 0,
      points: typeof user.points === "number" ? user.points : 0
    };
    this.userKey = encodeUserKey(username);
    return true;
  }

  async start(user, initialStatus = PRESENCE_STATES.HOME, meta = {}) {
    if (!this.setPublicUser(user)) return false;
    this.status = VALID_STATE_SET.has(initialStatus) ? initialStatus : PRESENCE_STATES.HOME;
    this.updateMeta(meta);
    this.ensureClockSync();

    if (this.active) {
      return this.writePresence(true);
    }

    this.active = true;
    const currentToken = ++this.token;

    try {
      this.connUnsub = onValue(ref(this.db, ".info/connected"), (snap) => {
        this.firebaseConnected = snap.val() === true;
        if (this.firebaseConnected && this.active && currentToken === this.token) {
          this.bindDisconnectAndPublish(currentToken);
        }
      }, () => {});
    } catch (e) {}

    clearInterval(this.heartbeatTimer);
    this.heartbeatTimer = setInterval(() => {
      if (this.active && currentToken === this.token && this.firebaseConnected) {
        this.writePresence(false);
      }
    }, PRESENCE_CONFIG.HEARTBEAT_INTERVAL_MS);

    if (typeof window !== "undefined" && typeof document !== "undefined") {
      this.unbindWake();
      this.wakeHandler = () => {
        if (this.active && currentToken === this.token && document.visibilityState !== "hidden") {
          this.writePresence(true);
        }
      };
      document.addEventListener("visibilitychange", this.wakeHandler);
      window.addEventListener("pageshow", this.wakeHandler);
      window.addEventListener("online", this.wakeHandler);
    }

    this.bindDisconnectAndPublish(currentToken);
    return true;
  }

  unbindWake() {
    if (!this.wakeHandler || typeof window === "undefined") return;
    document.removeEventListener("visibilitychange", this.wakeHandler);
    window.removeEventListener("pageshow", this.wakeHandler);
    window.removeEventListener("online", this.wakeHandler);
    this.wakeHandler = null;
  }

  updateMeta(meta = {}) {
    if (!meta || typeof meta !== "object") return;
    if ("roomId" in meta) this.meta.roomId = meta.roomId || null;
    if ("roomRole" in meta) this.meta.roomRole = meta.roomRole || null;
    if ("roomPlayerCount" in meta) this.meta.roomPlayerCount = typeof meta.roomPlayerCount === "number" ? meta.roomPlayerCount : 0;
    if ("hasClientInRoom" in meta) this.meta.hasClientInRoom = Boolean(meta.hasClientInRoom);
    if ("roomStatus" in meta) this.meta.roomStatus = meta.roomStatus || null;
  }

  async bindDisconnectAndPublish(token) {
    if (!this.active || token !== this.token || !this.userKey) return;
    const pRef = ref(this.db, `${GAME_ID}/presence/${this.userKey}`);
    try {
      const od = onDisconnect(pRef);
      await od.update({
        username: this.user.username,
        userKey: this.userKey,
        status: PRESENCE_STATES.OFFLINE,
        online: false,
        connected: false,
        sessionId: this.sessionId,
        disconnectedAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      this.onDisconnectRef = od;
    } catch (e) {}
    if (this.active && token === this.token) {
      await this.writePresence(true);
    }
  }

  async writePresence(claimOwnership = false) {
    if (!this.active || !this.userKey || !this.user) return false;
    const pRef = ref(this.db, `${GAME_ID}/presence/${this.userKey}`);
    try {
      const res = await runTransaction(pRef, (node) => {
        const now = this.getServerNow();
        // ป้องกันแท็บเก่าเขียนทับแท็บใหม่ที่กำลังออนไลน์อยู่
        if (!claimOwnership && node && node.sessionId && node.sessionId !== this.sessionId) {
          const hb = typeof node.heartbeatAt === "number" ? node.heartbeatAt : 0;
          if (node.online && (now - hb) < PRESENCE_CONFIG.HEARTBEAT_STALE_MS) {
            return;
          }
        }
        return {
          username: this.user.username,
          userKey: this.userKey,
          status: this.status,
          online: true,
          connected: true,
          sessionId: this.sessionId,
          roomId: this.meta.roomId,
          roomRole: this.meta.roomRole,
          roomPlayerCount: this.meta.roomPlayerCount,
          hasClientInRoom: this.meta.hasClientInRoom,
          roomStatus: this.meta.roomStatus,
          checkpoint: this.user.checkpoint,
          roundIndex: this.user.roundIndex,
          wins: this.user.wins,
          points: this.user.points,
          heartbeatAt: now,
          updatedAt: now,
          disconnectedAt: null
        };
      }, { applyLocally: false });
      return Boolean(res.committed);
    } catch (e) {
      return false;
    }
  }

  async updateStatus(status, meta = {}) {
    if (VALID_STATE_SET.has(status)) {
      this.status = status;
    }
    this.updateMeta(meta);
    if (!this.active) return false;
    return this.writePresence(false);
  }

  async stop(markOffline = true) {
    const targetKey = this.userKey;
    const targetSession = this.sessionId;
    this.token++;
    this.active = false;

    clearInterval(this.heartbeatTimer);
    this.heartbeatTimer = null;
    this.unbindWake();

    if (this.connUnsub) {
      try { this.connUnsub(); } catch (e) {}
      this.connUnsub = null;
    }
    if (this.onDisconnectRef) {
      try { await this.onDisconnectRef.cancel(); } catch (e) {}
      this.onDisconnectRef = null;
    }

    if (markOffline && targetKey) {
      const pRef = ref(this.db, `${GAME_ID}/presence/${targetKey}`);
      try {
        await runTransaction(pRef, (node) => {
          if (!node || (node.sessionId && node.sessionId !== targetSession)) return;
          const now = this.getServerNow();
          node.status = PRESENCE_STATES.OFFLINE;
          node.online = false;
          node.connected = false;
          node.roomId = null;
          node.roomRole = null;
          node.roomPlayerCount = 0;
          node.hasClientInRoom = false;
          node.roomStatus = null;
          node.disconnectedAt = now;
          node.updatedAt = now;
          return node;
        }, { applyLocally: false });
      } catch (e) {}
    }
  }

  async getPlayerPresence(usernameOrKey, inviter = null) {
    if (!usernameOrKey) return evaluatePresenceRecord(null);
    this.ensureClockSync();
    const key = encodeUserKey(String(usernameOrKey));
    try {
      const snap = await get(ref(this.db, `${GAME_ID}/presence/${key}`));
      const evaluated = evaluatePresenceRecord(
        snap.exists() ? snap.val() : null,
        this.getServerNow(),
        inviter || this.userKey
      );
      if (evaluated.online && evaluated.roomId) {
        await this.verifyWithLiveRoom(evaluated);
      }
      return evaluated;
    } catch (e) {
      return evaluatePresenceRecord(null);
    }
  }

  async verifyWithLiveRoom(evaluated) {
    try {
      const rSnap = await get(ref(this.db, `${GAME_ID}/onlineRooms/${evaluated.roomId}`));
      if (!rSnap.exists()) {
        evaluated.roomId = null;
        evaluated.hasClientInRoom = false;
        evaluated.roomPlayerCount = 0;
        if (evaluated.status === PRESENCE_STATES.WAITING_ROOM) {
          evaluated.status = PRESENCE_STATES.ONLINE_LOBBY;
          evaluated.isInvitable = true;
          evaluated.reason = "สามารถเชิญได้";
        }
        return;
      }
      const room = rSnap.val();
      const hasClient = Boolean(room.client && room.client.userKey);
      evaluated.roomStatus = room.status || null;
      evaluated.hasClientInRoom = hasClient;
      evaluated.roomPlayerCount = room.playerCount || (hasClient ? 2 : 1);

      if (room.status === "playing") {
        evaluated.status = PRESENCE_STATES.PLAYING_ONLINE;
        evaluated.isInvitable = false;
        evaluated.reason = "กำลังแข่งขันออนไลน์กับผู้เล่นจริง";
      } else if (room.status === "ready" || hasClient) {
        evaluated.status = PRESENCE_STATES.ONLINE_READY;
        evaluated.isInvitable = false;
        evaluated.reason = "อยู่ในห้องออนไลน์ที่มีคู่แข่งขันแล้ว";
      }
    } catch (e) {}
  }

  filterPlayers(rawMap = {}, options = {}) {
    const now = this.getServerNow();
    const excludeKey = options.excludeUsername
      ? encodeUserKey(String(options.excludeUsername))
      : this.userKey;
    const onlyInvitable = Boolean(options.onlyInvitable);

    return Object.values(rawMap || {})
      .map((rec) => evaluatePresenceRecord(rec, now, excludeKey))
      .filter((p) => {
        if (!p.online || p.isExpired) return false;
        if (excludeKey && p.userKey === excludeKey) return false;
        if (onlyInvitable && !p.isInvitable) return false;
        return true;
      })
      .sort((a, b) => {
        if (a.isInvitable !== b.isInvitable) return a.isInvitable ? -1 : 1;
        return (b.points || 0) - (a.points || 0);
      });
  }

  async getOnlinePlayers(options = {}) {
    this.ensureClockSync();
    try {
      const snap = await get(ref(this.db, `${GAME_ID}/presence`));
      return this.filterPlayers(snap.exists() ? snap.val() : {}, options);
    } catch (e) {
      return [];
    }
  }

  subscribe(callback, options = {}) {
    if (typeof callback !== "function") return () => {};
    this.ensureClockSync();

    const id = this.subSeq++;
    this.subscribers.set(id, { callback, options });

    if (!this.listUnsub) {
      this.listUnsub = onValue(ref(this.db, `${GAME_ID}/presence`), (snap) => {
        this.rawCache = snap.exists() ? (snap.val() || {}) : {};
        this.emitToSubscribers();
      }, () => {});

      clearInterval(this.sweepTimer);
      this.sweepTimer = setInterval(() => this.emitToSubscribers(), PRESENCE_CONFIG.SWEEP_INTERVAL_MS);
    } else {
      this.emitSingle(id);
    }

    return () => {
      this.subscribers.delete(id);
      if (this.subscribers.size === 0) {
        if (this.listUnsub) {
          try { this.listUnsub(); } catch (e) {}
          this.listUnsub = null;
        }
        clearInterval(this.sweepTimer);
        this.sweepTimer = null;
        this.rawCache = {};
      }
    };
  }

  emitSingle(id) {
    const sub = this.subscribers.get(id);
    if (!sub) return;
    try {
      const allOnline = this.filterPlayers(this.rawCache, {
        onlyInvitable: false,
        excludeUsername: sub.options.excludeUsername || this.userKey
      });
      const invitable = allOnline.filter((p) => p.isInvitable);
      sub.callback(sub.options.onlyInvitable ? invitable : allOnline, {
        allOnline,
        invitable,
        serverTimeMs: this.getServerNow()
      });
    } catch (e) {}
  }

  emitToSubscribers() {
    for (const id of this.subscribers.keys()) {
      this.emitSingle(id);
    }
  }

  async verifyCanInvite(targetUsername, inviterUsername = null) {
    const p = await this.getPlayerPresence(targetUsername, inviterUsername || this.userKey);
    return {
      canInvite: Boolean(p.online && !p.isExpired && p.isInvitable),
      reason: p.reason,
      presence: p
    };
  }

  async verifyCanAcceptInvitation({ inviteeUsername, hostUsername, roomId } = {}) {
    const inviteeKey = encodeUserKey(String(inviteeUsername || this.userKey || ""));
    const hostKey = hostUsername ? encodeUserKey(String(hostUsername)) : null;

    if (!inviteeKey) {
      return { allowed: false, reason: "กรุณาเข้าสู่ระบบก่อนตอบรับคำเชิญ" };
    }
    if (hostKey && inviteeKey === hostKey) {
      return { allowed: false, reason: "ไม่สามารถตอบรับคำเชิญของตัวเองได้" };
    }

    const invitee = await this.getPlayerPresence(inviteeKey, null);
    if (!invitee.online || invitee.isExpired) {
      return { allowed: false, reason: "สถานะออนไลน์ของคุณหมดอายุ กรุณารีเฟรชหน้าเว็บ", invitee };
    }
    if (invitee.status === PRESENCE_STATES.PLAYING_ONLINE || invitee.status === PRESENCE_STATES.ONLINE_READY) {
      return { allowed: false, reason: "คุณกำลังอยู่ในห้องที่มีคู่แข่งหรือกำลังแข่งขันอยู่", invitee };
    }

    let host = null;
    if (hostKey) {
      host = await this.getPlayerPresence(hostKey, null);
      if (!host.online || host.isExpired) {
        return { allowed: false, reason: "ผู้เชิญออฟไลน์หรือหลุดการเชื่อมต่อแล้ว", invitee, host };
      }
      if (host.status === PRESENCE_STATES.PLAYING_ONLINE) {
        return { allowed: false, reason: "ผู้เชิญกำลังแข่งขันกับผู้เล่นอื่นอยู่", invitee, host };
      }
    }

    if (roomId) {
      try {
        const rSnap = await get(ref(this.db, `${GAME_ID}/onlineRooms/${roomId}`));
        if (!rSnap.exists()) {
          return { allowed: false, reason: "ไม่พบห้องแข่งขันนี้ ห้องอาจถูกปิดไปแล้ว", invitee, host };
        }
        const room = rSnap.val();
        if (room.status !== "waiting") {
          return { allowed: false, reason: "ห้องนี้ไม่ได้อยู่ในสถานะรอผู้เล่นแล้ว", invitee, host, room };
        }
        if (room.client && room.client.userKey && room.client.userKey !== inviteeKey) {
          return { allowed: false, reason: "ห้องนี้มีผู้เล่นเข้าร่วมเต็มแล้ว", invitee, host, room };
        }
        if (!room.host || room.host.connected === false) {
          return { allowed: false, reason: "หัวหน้าห้องหลุดการเชื่อมต่อ", invitee, host, room };
        }
        return { allowed: true, reason: "สามารถเข้าร่วมห้องได้", invitee, host, room };
      } catch (e) {
        return { allowed: false, reason: "ไม่สามารถตรวจสอบสถานะห้องล่าสุดได้", invitee, host };
      }
    }

    if (host && (host.status === PRESENCE_STATES.ONLINE_READY || host.hasClientInRoom)) {
      return { allowed: false, reason: "ห้องของผู้เชิญมีผู้เล่นอื่นเข้าร่วมแล้ว", invitee, host };
    }

    return { allowed: true, reason: "สามารถตอบรับคำเชิญได้", invitee, host };
  }

  async cleanupAll(markOffline = true) {
    this.subscribers.clear();
    if (this.listUnsub) {
      try { this.listUnsub(); } catch (e) {}
      this.listUnsub = null;
    }
    clearInterval(this.sweepTimer);
    this.sweepTimer = null;
    await this.stop(markOffline);
  }
}

export const presenceManager = new PlayerPresenceService(database);

// Export Functions สำหรับให้ ping.html และ ping-online.js เรียกใช้งานโดยตรง
export const startPlayerPresence = (user, status, meta) => presenceManager.start(user, status, meta);
export const updatePlayerPresence = (status, meta) => presenceManager.updateStatus(status, meta);
export const getPlayerPresence = (username, inviter) => presenceManager.getPlayerPresence(username, inviter);
export const getOnlinePlayers = (options) => presenceManager.getOnlinePlayers(options);
export const getInvitablePlayers = (excludeUsername) => presenceManager.getOnlinePlayers({ onlyInvitable: true, excludeUsername });
export const subscribeOnlinePlayers = (callback, options) => presenceManager.subscribe(callback, options);
export const filterInvitablePlayers = (rawMap, excludeUsername) => presenceManager.filterPlayers(rawMap, { onlyInvitable: true, excludeUsername });
export const verifyPlayerCanBeInvited = (target, inviter) => presenceManager.verifyCanInvite(target, inviter);
export const verifyCanAcceptInvitation = (options) => presenceManager.verifyCanAcceptInvitation(options);
export const stopPlayerPresence = (markOffline = true) => presenceManager.cleanupAll(markOffline);

/**
 * =========================================================================
 * AuthManager (คงการทำงานและเมธอดเดิมทั้งหมด 100%)
 * =========================================================================
 */
export class AuthManager {
  constructor() {
    this.currentUser = null;
    this.db = database;
    this.presence = presenceManager;
  }

  async findUser(username) {
    const key = encodeUserKey(username);
    const snap = await get(ref(this.db, `${GAME_ID}/players/${key}`));
    if (snap.exists()) {
      return snap.val();
    }
    return null;
  }

  // เรียกใช้อัตโนมัติเมื่อเปิดหน้าเว็บ เพื่อดึงเซสชันเดิมกลับมา
  async initAuthStateListener(onUserAuthenticated, onUnauthenticated) {
    const key = SessionStorage.getSavedKey();
    if (key) {
      try {
        const snap = await get(ref(this.db, `${GAME_ID}/players/${key}`));
        if (snap.exists()) {
          const profile = snap.val();
          if (typeof profile.points !== 'number') profile.points = 0;
          if (typeof profile.unlockedRoundIndex !== 'number') {
            profile.unlockedRoundIndex = profile.roundIndex || 0;
          }
          this.currentUser = profile;
          SessionStorage.save(profile);
          this.presence.start(profile, PRESENCE_STATES.HOME).catch(() => {});
          if (typeof onUserAuthenticated === 'function') {
            onUserAuthenticated(profile);
          }
          return;
        }
      } catch (err) {
        console.warn("ไม่สามารถกู้คืนเซสชันได้:", err);
      }
    }
    SessionStorage.clear();
    if (typeof onUnauthenticated === 'function') {
      onUnauthenticated();
    }
  }

  async login(username, password) {
    const user = await this.findUser(username);
    if (!user) {
      const err = new Error("ไม่พบบัญชีผู้เล่นนี้ในระบบ กรุณาตรวจสอบชื่อหรือสมัครสมาชิกใหม่");
      err.code = "auth/user-not-found";
      throw err;
    }
    if (user.password !== password) {
      const err = new Error("รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง");
      err.code = "auth/wrong-password";
      throw err;
    }
    if (typeof user.points !== 'number') user.points = 0;
    if (typeof user.unlockedRoundIndex !== 'number') {
      user.unlockedRoundIndex = user.roundIndex || 0;
    }
    this.currentUser = user;
    SessionStorage.save(user);
    this.presence.start(user, PRESENCE_STATES.HOME).catch(() => {});
    return user;
  }

  async register(username, password) {
    const existing = await this.findUser(username);
    if (existing) {
      const err = new Error("ชื่อผู้เล่นนี้มีอยู่ในระบบแล้ว กรุณาเลือกชื่ออื่น หรือเข้าสู่ระบบ");
      err.code = "auth/username-already-exists";
      throw err;
    }
    const newUser = {
      username: username.trim(),
      password: password,
      checkpoint: "รอบ 16 คนสุดท้าย",
      roundIndex: 0,
      unlockedRoundIndex: 0,
      wins: 0,
      points: 0,
      updatedAt: Date.now()
    };
    const key = encodeUserKey(username);
    await set(ref(this.db, `${GAME_ID}/players/${key}`), newUser);
    this.currentUser = newUser;
    SessionStorage.save(newUser);
    this.presence.start(newUser, PRESENCE_STATES.HOME).catch(() => {});
    return newUser;
  }

  async saveProgress(roundIndex, checkpointName, incrementWin = false, earnedPoints = 0) {
    if (!this.currentUser || !this.currentUser.username) return null;
    const key = encodeUserKey(this.currentUser.username);
    this.currentUser.roundIndex = roundIndex;
    this.currentUser.checkpoint = checkpointName;

    const currentUnlocked = typeof this.currentUser.unlockedRoundIndex === 'number'
      ? this.currentUser.unlockedRoundIndex
      : 0;
    this.currentUser.unlockedRoundIndex = Math.max(currentUnlocked, roundIndex);

    if (incrementWin) {
      this.currentUser.wins = (this.currentUser.wins || 0) + 1;
    }
    if (earnedPoints > 0) {
      this.currentUser.points = (this.currentUser.points || 0) + earnedPoints;
    }
    this.currentUser.updatedAt = Date.now();

    await update(ref(this.db, `${GAME_ID}/players/${key}`), {
      roundIndex: this.currentUser.roundIndex,
      unlockedRoundIndex: this.currentUser.unlockedRoundIndex,
      checkpoint: this.currentUser.checkpoint,
      wins: this.currentUser.wins || 0,
      points: this.currentUser.points || 0,
      updatedAt: this.currentUser.updatedAt
    });

    SessionStorage.save(this.currentUser);
    if (this.presence.active) {
      this.presence.setPublicUser(this.currentUser);
      this.presence.writePresence(false).catch(() => {});
    }
    return this.currentUser;
  }

  async getOnlineStats(username) {
    const targetName = username || (this.currentUser ? this.currentUser.username : null);
    if (!targetName) return null;
    const key = encodeUserKey(targetName);
    try {
      const snap = await get(ref(this.db, `${GAME_ID}/onlineStats/${key}`));
      if (snap.exists()) {
        return snap.val();
      }
    } catch (e) {
      console.warn("ไม่สามารถดึงสถิติออนไลน์ได้:", e);
    }
    return {
      username: targetName,
      totalMatches: 0,
      wins: 0,
      losses: 0,
      aborted: 0,
      lastPlayedAt: null
    };
  }

  async getLeaderboard() {
    try {
      const snap = await get(ref(this.db, `${GAME_ID}/players`));
      if (snap.exists()) {
        const list = Object.values(snap.val());
        return list.sort((a, b) => {
          if ((b.points || 0) !== (a.points || 0)) return (b.points || 0) - (a.points || 0);
          if ((b.wins || 0) !== (a.wins || 0)) return (b.wins || 0) - (a.wins || 0);
          return (b.roundIndex || 0) - (a.roundIndex || 0);
        }).slice(0, 20);
      }
    } catch (e) {
      console.error("ดึง Leaderboard ไม่สำเร็จ:", e);
    }
    return [];
  }

  // Method สำหรับเรียกใช้ระบบ Presence ผ่าน authManager
  updatePresenceStatus(status, meta = {}) {
    return this.presence.updateStatus(status, meta);
  }

  getPlayerPresence(username) {
    return this.presence.getPlayerPresence(username, this.currentUser ? this.currentUser.username : null);
  }

  getInvitablePlayers() {
    return this.presence.getOnlinePlayers({
      onlyInvitable: true,
      excludeUsername: this.currentUser ? this.currentUser.username : null
    });
  }

  subscribeOnlinePlayers(callback, options = {}) {
    return this.presence.subscribe(callback, options);
  }

  verifyPlayerCanBeInvited(targetUsername) {
    return this.presence.verifyCanInvite(targetUsername, this.currentUser ? this.currentUser.username : null);
  }

  verifyCanAcceptInvitation(options = {}) {
    return this.presence.verifyCanAcceptInvitation({
      ...options,
      inviteeUsername: options.inviteeUsername || (this.currentUser ? this.currentUser.username : null)
    });
  }

  async logout() {
    await this.presence.cleanupAll(true);
    this.currentUser = null;
    SessionStorage.clear();
  }
}