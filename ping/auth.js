/**
 * auth.js
 * ระบบยืนยันตัวตนและจัดการเซสชันผู้เล่น (PingPong Realtime Database + Session Persistence)
 * รองรับการจำสถานะข้าม Safari iOS / In-App Browser และทนต่อการกด Refresh 100%
 * พร้อมระบบ Presence ติดตามสถานะผู้เล่นออนไลน์และกิจกรรมปัจจุบันแบบ Realtime
 */
import { initializeApp, getApps, getApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { 
  getDatabase, 
  ref, 
  get, 
  set, 
  update,
  remove,
  onValue,
  onDisconnect
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

export class AuthManager {
  constructor() {
    this.currentUser = null;
    this.db = database;
    this.currentActivity = 'home';
    this.currentRoomId = null;
    this.presenceUnsubscribe = null;
    this.presenceDisconnectRef = null;
  }

  async findUser(username) {
    const key = encodeUserKey(username);
    const snap = await get(ref(this.db, `${GAME_ID}/players/${key}`));
    if (snap.exists()) {
      return snap.val();
    }
    return null;
  }

  /**
   * เริ่มต้นระบบติดตามสถานะออนไลน์ (Presence) ของผู้เล่นที่ล็อกอินอยู่
   * เมื่อปิดแท็บหรือเน็ตหลุด Firebase จะลบข้อมูลออกจากโหนด presence โดยอัตโนมัติ
   */
  setupPresence(username) {
    if (!username) return;
    this.cleanupPresenceListenerOnly();

    const userKey = encodeUserKey(username);
    const presenceRef = ref(this.db, `${GAME_ID}/presence/${userKey}`);
    const connectedRef = ref(this.db, ".info/connected");

    this.presenceUnsubscribe = onValue(connectedRef, async (snap) => {
      if (snap.val() === true && this.currentUser) {
        try {
          this.presenceDisconnectRef = onDisconnect(presenceRef);
          await this.presenceDisconnectRef.remove();
          await set(presenceRef, {
            username: this.currentUser.username,
            userKey: userKey,
            online: true,
            activity: this.currentActivity || 'home',
            roomId: this.currentRoomId || null,
            updatedAt: Date.now()
          });
        } catch (err) {
          console.warn("ไม่สามารถตั้งค่าสถานะออนไลน์ได้:", err);
        }
      }
    });
  }

  /**
   * อัปเดตสถานะกิจกรรมปัจจุบันของผู้เล่น เช่น
   * 'home' | 'bot_match' | 'lobby' | 'waiting_room' | 'online_match'
   */
  async updateActivity(activity, roomId = null) {
    this.currentActivity = activity;
    this.currentRoomId = roomId;
    if (!this.currentUser || !this.currentUser.username) return;

    const userKey = encodeUserKey(this.currentUser.username);
    const presenceRef = ref(this.db, `${GAME_ID}/presence/${userKey}`);
    try {
      await update(presenceRef, {
        username: this.currentUser.username,
        userKey: userKey,
        online: true,
        activity: activity,
        roomId: roomId,
        updatedAt: Date.now()
      });
    } catch (err) {
      console.warn("ไม่สามารถอัปเดตสถานะกิจกรรมได้:", err);
    }
  }

  cleanupPresenceListenerOnly() {
    if (this.presenceUnsubscribe) {
      this.presenceUnsubscribe();
      this.presenceUnsubscribe = null;
    }
  }

  async clearPresence() {
    this.cleanupPresenceListenerOnly();
    if (this.presenceDisconnectRef) {
      try {
        await this.presenceDisconnectRef.cancel();
      } catch (e) {}
      this.presenceDisconnectRef = null;
    }
    if (this.currentUser && this.currentUser.username) {
      const userKey = encodeUserKey(this.currentUser.username);
      try {
        await remove(ref(this.db, `${GAME_ID}/presence/${userKey}`));
      } catch (e) {}
    }
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
          this.setupPresence(profile.username);
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
    this.setupPresence(user.username);
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
    this.setupPresence(newUser.username);
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

  async logout() {
    await this.clearPresence();
    this.currentUser = null;
    SessionStorage.clear();
  }
}