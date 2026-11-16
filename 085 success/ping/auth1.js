/**
 * auth.js
 * ระบบยืนยันตัวตนและจัดการเซสชันผู้เล่น (PingPong Realtime Database + Session Persistence)
 * รองรับการจำสถานะข้าม Safari iOS / In-App Browser และทนต่อการกด Refresh 100%
 */
import { initializeApp, getApps, getApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { 
  getDatabase, 
  ref, 
  get, 
  set, 
  update 
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
const GAME_ID = "PingPong";

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
          this.currentUser = profile;
          SessionStorage.save(profile);
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
    this.currentUser = user;
    SessionStorage.save(user);
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
      checkpoint: "รอบ 8 คนสุดท้าย",
      roundIndex: 0,
      wins: 0,
      updatedAt: Date.now()
    };
    const key = encodeUserKey(username);
    await set(ref(this.db, `${GAME_ID}/players/${key}`), newUser);
    this.currentUser = newUser;
    SessionStorage.save(newUser);
    return newUser;
  }

  async saveProgress(roundIndex, checkpointName, incrementWin = false) {
    if (!this.currentUser || !this.currentUser.username) return null;
    const key = encodeUserKey(this.currentUser.username);
    this.currentUser.roundIndex = roundIndex;
    this.currentUser.checkpoint = checkpointName;
    if (incrementWin) {
      this.currentUser.wins = (this.currentUser.wins || 0) + 1;
    }
    this.currentUser.updatedAt = Date.now();

    await update(ref(this.db, `${GAME_ID}/players/${key}`), {
      roundIndex: this.currentUser.roundIndex,
      checkpoint: this.currentUser.checkpoint,
      wins: this.currentUser.wins,
      updatedAt: this.currentUser.updatedAt
    });

    SessionStorage.save(this.currentUser);
    return this.currentUser;
  }

  async getLeaderboard() {
    try {
      const snap = await get(ref(this.db, `${GAME_ID}/players`));
      if (snap.exists()) {
        const list = Object.values(snap.val());
        return list.sort((a, b) => {
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
    this.currentUser = null;
    SessionStorage.clear();
  }
}