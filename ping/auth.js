/**
 * auth.js
 * ระบบจัดการ Authentication และซิงค์สถานะผู้เล่นสำหรับ Blind Ping Pong 3D
 * รองรับการจำ Session ข้ามแท็บ/Refresh บน Safari iOS และ Desktop
 */

import { initializeApp, getApps, getApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { 
  getAuth, 
  onAuthStateChanged, 
  signInWithEmailAndPassword, 
  createUserWithEmailAndPassword, 
  signOut,
  setPersistence,
  browserLocalPersistence 
} from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";
import { 
  getDatabase, 
  ref, 
  get, 
  set, 
  child 
} from "https://www.gstatic.com/firebasejs/10.8.0/firebase-database.js";

// 1. ตั้งค่า Firebase Project
const firebaseConfig = {
  apiKey: "AIzaSyDvcdgsyT5sDdYTYKIqetzNL9Be-MFC0l4",
  authDomain: "xo-game-134ec.firebaseapp.com",
  databaseURL: "https://xo-game-134ec-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "xo-game-134ec",
  storageBucket: "xo-game-134ec.firebasestorage.app",
  messagingSenderId: "318375224157",
  appId: "1:318375224157:web:35a997751e82f7b6b77eb4"
};

// ป้องกันการ Initialize ซ้ำ
const app = !getApps().length ? initializeApp(firebaseConfig) : getApp();
export const auth = getAuth(app);
export const database = getDatabase(app);

const GAME_ID = "PingPong";

/**
 * แปลง Username ทั่วไปเป็น Dummy Email เพื่อใช้งานร่วมกับ Firebase Auth
 * (เนื่องจาก Firebase Auth บังคับรูปแบบ Email)
 */
function usernameToEmail(username) {
  const cleanName = encodeURIComponent(username.trim().toLowerCase()).replace(/%/g, 'x');
  return `${cleanName}@pingpong.local`;
}

/**
 * เข้ารหัส Username ให้ปลอดภัยสำหรับใช้เป็น Node Key ใน Realtime Database
 */
export function encodeUserKey(username) {
  return encodeURIComponent(username.trim().toLowerCase()).replace(/\./g, '%2E');
}

/**
 * ตั้งค่า Local Persistence ให้จำรหัสผ่าน/เซสชันถาวรบนอุปกรณ์
 */
setPersistence(auth, browserLocalPersistence).catch((err) => {
  console.warn("ไม่สามารถเปิดใช้งาน Local Persistence:", err);
});

/**
 * ตรวจสอบความถูกต้องของชื่อและรหัสผ่าน
 */
const BANNED_WORDS = [
  'ควย', 'เหี้ย', 'เย็ด', 'มึง', 'กู', 'สัส', 'ห่า', 'ดอกทอง', 'ระยำ', 'จัญไร',
  'ชาติชั่ว', 'สันดาน', 'หน้าด้าน', 'กะหรี่', 'อีสัตว์', 'ไอ้สัตว์', 'แม่เย็ด',
  'fuck', 'shit', 'bitch', 'asshole', 'dick', 'bastard', 'cunt', 'pussy'
];

export function validateUserData(username, password) {
  if (!username || !password) return { valid: false, msg: "กรุณากรอกข้อมูลให้ครบถ้วน" };
  if (username.length > 20) return { valid: false, msg: "ชื่อผู้เล่นต้องไม่เกิน 20 ตัวอักษร" };
  if (password.length < 6) return { valid: false, msg: "รหัสผ่านต้องมีความยาวอย่างน้อย 6 ตัวอักษร" }; // Firebase Auth กำหนด 6 ตัวขึ้นไป
  
  const clean = username.toLowerCase().replace(/\s+/g, '');
  if (BANNED_WORDS.some(word => clean.includes(word))) {
    return { valid: false, msg: "ชื่อผู้เล่นมีคำไม่สุภาพ กรุณาเปลี่ยนชื่อใหม่" };
  }
  return { valid: true };
}

/**
 * สมัครสมาชิกใหม่: ทำงานทั้งบน Firebase Auth และ Realtime Database
 */
export async function registerPlayer(username, password) {
  const check = validateUserData(username, password);
  if (!check.valid) throw new Error(check.msg);

  const email = usernameToEmail(username);
  const userCredential = await createUserWithEmailAndPassword(auth, email, password);
  const uid = userCredential.user.uid;

  const initialProfile = {
    uid: uid,
    username: username.trim(),
    checkpoint: "รอบ 8 คนสุดท้าย",
    roundIndex: 0,
    wins: 0,
    updatedAt: Date.now()
  };

  const key = encodeUserKey(username);
  await set(ref(database, `${GAME_ID}/players/${key}`), initialProfile);
  // ผูก UID Map เข้ากับ Key เพื่อความแม่นยำในการดึงข้อมูล
  await set(ref(database, `${GAME_ID}/uid_map/${uid}`), key);

  return initialProfile;
}

/**
 * เข้าสู่ระบบ
 */
export async function loginPlayer(username, password) {
  const email = usernameToEmail(username);
  const userCredential = await signInWithEmailAndPassword(auth, email, password);
  const userProfile = await fetchProfileByUid(userCredential.user.uid);
  return userProfile;
}

/**
 * ออกจากระบบ
 */
export async function logoutPlayer() {
  await signOut(auth);
}

/**
 * ดึงโปรไฟล์ผู้เล่นตาม UID ปัจจุบัน
 */
export async function fetchProfileByUid(uid) {
  const keySnapshot = await get(ref(database, `${GAME_ID}/uid_map/${uid}`));
  if (keySnapshot.exists()) {
    const key = keySnapshot.val();
    const userSnapshot = await get(ref(database, `${GAME_ID}/players/${key}`));
    if (userSnapshot.exists()) {
      return userSnapshot.val();
    }
  }
  return null;
}

/**
 * อัปเดตความคืบหน้าด่าน/ชัยชนะ
 */
export async function updatePlayerProgress(username, roundIndex, checkpointName, incrementWin = false) {
  const key = encodeUserKey(username);
  const playerRef = ref(database, `${GAME_ID}/players/${key}`);
  const snapshot = await get(playerRef);

  if (snapshot.exists()) {
    const user = snapshot.val();
    user.roundIndex = roundIndex;
    user.checkpoint = checkpointName;
    if (incrementWin) {
      user.wins = (user.wins || 0) + 1;
    }
    user.updatedAt = Date.now();
    await set(playerRef, user);
    return user;
  }
  return null;
}

/**
 * ตรวจสอบสถานะการล็อกอินอัตโนมัติ (หัวใจสำคัญแก้ปัญหา Safari Refresh/สลับหน้าต่าง)
 * @param {Function} onLoggedIn Callback เมื่อพบเซสชันเดิม
 * @param {Function} onLoggedOut Callback เมื่อไม่พบเซสชัน
 */
export function initAuthListener(onLoggedIn, onLoggedOut) {
  return onAuthStateChanged(auth, async (user) => {
    if (user) {
      try {
        const profile = await fetchProfileByUid(user.uid);
        if (profile) {
          onLoggedIn(profile);
          return;
        }
      } catch (err) {
        console.error("เกิดข้อผิดพลาดในการดึงข้อมูลโปรไฟล์:", err);
      }
    }
    onLoggedOut();
  });
}