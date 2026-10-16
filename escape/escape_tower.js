import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
import { getDatabase, ref, set, onValue, update, runTransaction, remove, onDisconnect } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-database.js";

// Firebase Config
const firebaseConfig = {
    apiKey: "AIzaSyDvcdgsyT5sDdYTYKIqetzNL9Be-MFC0l4",
    authDomain: "xo-game-134ec.firebaseapp.com",
    databaseURL: "https://xo-game-134ec-default-rtdb.asia-southeast1.firebasedatabase.app",
    projectId: "xo-game-134ec",
    storageBucket: "xo-game-134ec.firebasestorage.app",
    messagingSenderId: "318375224157",
    appId: "1:318375224157:web:667876ec774f9f03b77eb4"
};

const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

// Global State
let myName = '';
let myId = null;
let currentRoomId = null;
let isHost = false;
let gameState = null;
let roomListener = null;
let speechQueue = [];
let isSpeaking = false;
let lastAnnouncedTurn = null;
let activeTurnLock = null; // ป้องกันการซ้อนทับของ Callback แบบระบุ Turn Key

// Scene & BGM State
let bgmSource = null;
let isBgmPlaying = false;
let startCutscenePlayed = false;
let winCutscenePlayed = false;
let isPlayingStartCutscene = false;
let isPlayingWinCutscene = false;
let sceneTimeouts = [];

// Visual FX Tracking State (ป้องกันการแสดงเอฟเฟกต์ซ้ำจากเหตุการณ์เดียวกัน)
let lastVfxLogTs = 0;
let lastVfxTurnKey = null;

// Timer & Bot Names
let turnTimer = null;
let currentTurnKey = null;
const TURN_TIME_LIMIT = 55000;
const BOT_NAMES = [
    "ไข่เค็ม(b)", "ไข่ดาว(B)", "ไข่เจียว(B)", "ไข่ต้ม(B)", "ไข่คน(B)", 
    "ไข่ตุ๋น(B)", "ไข่กระทะ(B)", "ไข่ครอบ(B)", "ไข่ยัดไส้(B)", "ไข่ระเบิด(B)"
];

// ==================== Visual Assets & Helpers (SVG Data URIs - ไม่พึ่งพาไฟล์ภายนอก) ====================
const HUMAN_AVATARS = [
    "data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'><rect width='64' height='64' rx='32' fill='%230f172a'/><circle cx='32' cy='24' r='12' fill='%23fbbf24'/><path d='M14 56c0-11 8-18 18-18s18 7 18 18' fill='%233b82f6'/></svg>",
    "data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'><rect width='64' height='64' rx='32' fill='%230f172a'/><circle cx='32' cy='24' r='12' fill='%23f59e0b'/><path d='M14 56c0-11 8-18 18-18s18 7 18 18' fill='%2310b981'/></svg>",
    "data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'><rect width='64' height='64' rx='32' fill='%230f172a'/><circle cx='32' cy='24' r='12' fill='%23fde047'/><path d='M14 56c0-11 8-18 18-18s18 7 18 18' fill='%238b5cf6'/></svg>",
    "data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'><rect width='64' height='64' rx='32' fill='%230f172a'/><circle cx='32' cy='24' r='12' fill='%23fdba74'/><path d='M14 56c0-11 8-18 18-18s18 7 18 18' fill='%23ec4899'/></svg>"
];

const BOT_AVATARS = [
    "data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'><rect width='64' height='64' rx='32' fill='%230f172a'/><rect x='16' y='14' width='32' height='22' rx='5' fill='%2394a3b8'/><circle cx='24' cy='25' r='3.5' fill='%23ef4444'/><circle cx='40' cy='25' r='3.5' fill='%23ef4444'/><path d='M14 56c0-10 8-16 18-16s18 6 18 16' fill='%23475569'/></svg>",
    "data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'><rect width='64' height='64' rx='32' fill='%230f172a'/><rect x='16' y='14' width='32' height='22' rx='5' fill='%2364748b'/><circle cx='24' cy='25' r='3.5' fill='%23fbbf24'/><circle cx='40' cy='25' r='3.5' fill='%23fbbf24'/><path d='M14 56c0-10 8-16 18-16s18 6 18 16' fill='%23334155'/></svg>",
    "data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'><rect width='64' height='64' rx='32' fill='%230f172a'/><rect x='16' y='14' width='32' height='22' rx='5' fill='%23cbd5e1'/><circle cx='24' cy='25' r='3.5' fill='%2322c55e'/><circle cx='40' cy='25' r='3.5' fill='%2322c55e'/><path d='M14 56c0-10 8-16 18-16s18 6 18 16' fill='%231e293b'/></svg>"
];

// แผนผังสัญลักษณ์และ Class ของประเภทห้องตามระบบจริง (และชื่อเทียบเคียงตามข้อกำหนด)
const ROOM_VISUAL_MAP = {
    key: { icon: '🔑', cls: 'room-type-key' },
    key_normal: { icon: '🔑', cls: 'room-type-key' },
    ghostKey: { icon: '🗝️', cls: 'room-type-ghostKey' },
    key_skeleton: { icon: '🗝️', cls: 'room-type-ghostKey' },
    secret: { icon: '🪜', cls: 'room-type-secret' },
    secret_stair: { icon: '🪜', cls: 'room-type-secret' },
    zombie: { icon: '🧟', cls: 'room-type-zombie' },
    vaccine: { icon: '💉', cls: 'room-type-vaccine' },
    food: { icon: '🥫', cls: 'room-type-food' },
    armor: { icon: '🛡️', cls: 'room-type-armor' },
    puzzle: { icon: '🚽', cls: 'room-type-puzzle' },
    gun: { icon: '🔫', cls: 'room-type-gun' },
    laser_gun: { icon: '🔫', cls: 'room-type-gun' },
    battery: { icon: '🔋', cls: 'room-type-battery' }
};

function escapeHtml(str) {
    return String(str ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function isClueRoomType(type) {
    return type === 'clue' || type === 'hint_1' || type === 'hint_2' || type === 'hint_3';
}

function getParticipantAvatarSrc(pId, playerObj) {
    const slotNum = parseInt(String(pId || 'p1').replace(/\D/g, ''), 10) || 1;
    if (playerObj && playerObj.isBot) {
        const nameIdx = BOT_NAMES.indexOf(playerObj.name);
        const botIdx = nameIdx >= 0 ? (nameIdx % BOT_AVATARS.length) : ((slotNum - 1) % BOT_AVATARS.length);
        return BOT_AVATARS[botIdx];
    }
    const humanIdx = (slotNum - 1) % HUMAN_AVATARS.length;
    return HUMAN_AVATARS[humanIdx];
}

function getParticipantAvatarHtml(pId, playerObj) {
    const src = getParticipantAvatarSrc(pId, playerObj);
    const cls = (playerObj && playerObj.isBot) ? 'bot-avatar' : 'character-avatar';
    return `<img src="${src}" alt="" class="${cls}" aria-hidden="true">`;
}

function ensureVisualRuntimeStyles() {
    if (document.getElementById('escape-visual-runtime-style')) return;
    const style = document.createElement('style');
    style.id = 'escape-visual-runtime-style';
    style.textContent = `
        /* ปิด pseudo-element เดิมเมื่อมีรูป avatar จริง เพื่อไม่ให้แสดงไอคอนซ้ำ 2 อัน */
        #participant-list .participant-card.has-js-avatar::before,
        #all-status-list .participant-card.has-js-avatar::before,
        #result-list .participant-card.has-js-avatar::before {
            display: none !important;
            content: none !important;
        }
        /* สไตล์ไอคอนประกอบสถานะผู้เล่นปัจจุบันและแผงสถานะตัวละครทั้งหมด */
        .et-status-bar-items {
            display: inline-flex;
            flex-wrap: wrap;
            align-items: center;
            justify-content: center;
            gap: 6px 8px;
        }
        .et-stat-pill {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            padding: 2px 7px;
            border-radius: 6px;
            background: rgba(15, 23, 42, 0.75);
            border: 1px solid rgba(148, 163, 184, 0.25);
            white-space: nowrap;
            font-size: 0.92em;
            line-height: 1.3;
        }
        .et-stat-icon {
            display: inline-block;
            line-height: 1;
            filter: drop-shadow(0 1px 2px rgba(0, 0, 0, 0.6));
        }
        .et-all-status-body {
            display: flex;
            flex-direction: column;
            gap: 6px;
            width: 100%;
        }
        .et-all-status-header {
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 6px 10px;
            font-weight: 700;
            color: #f8fafc;
        }
        .et-all-status-grid {
            display: flex;
            flex-wrap: wrap;
            gap: 5px 7px;
        }
        .et-state-tag {
            display: inline-flex;
            align-items: center;
            gap: 3px;
            padding: 1px 7px;
            border-radius: 999px;
            font-size: 0.8rem;
            font-weight: 600;
            background: rgba(239, 68, 68, 0.2);
            border: 1px solid rgba(239, 68, 68, 0.5);
            color: #fca5a5;
        }
        /* การแสดงผลห้องที่สำรวจแล้วตามผลลัพธ์จริง (ยกเว้นห้องคำใบ้ที่คงเป็นประตูปิด) */
        .board-cell.explored:not(.room-clue-door) {
            text-decoration: none !important;
            opacity: 0.92 !important;
            gap: 2px;
        }
        .board-cell .room-num-label {
            font-size: 0.72em;
            line-height: 1;
            opacity: 0.85;
        }
        .board-cell .room-result-icon {
            font-size: 1.05em;
            line-height: 1;
            filter: drop-shadow(0 2px 3px rgba(0,0,0,0.85));
        }
        .board-cell.explored.room-type-zombie {
            border-color: #ef4444 !important;
            background: radial-gradient(circle at 50% 50%, rgba(185, 28, 28, 0.38) 0%, #0b0f17 85%) !important;
            color: #fca5a5 !important;
        }
        .board-cell.explored.room-type-key,
        .board-cell.explored.room-type-ghostKey,
        .board-cell.explored.room-type-secret {
            border-color: #fbbf24 !important;
            background: radial-gradient(circle at 50% 50%, rgba(245, 158, 11, 0.32) 0%, #0b0f17 85%) !important;
            color: #fde68a !important;
        }
        .board-cell.explored.room-type-vaccine,
        .board-cell.explored.room-type-food {
            border-color: #22c55e !important;
            background: radial-gradient(circle at 50% 50%, rgba(22, 163, 74, 0.32) 0%, #0b0f17 85%) !important;
            color: #bbf7d0 !important;
        }
        .board-cell.explored.room-type-armor,
        .board-cell.explored.room-type-gun,
        .board-cell.explored.room-type-battery {
            border-color: #38bdf8 !important;
            background: radial-gradient(circle at 50% 50%, rgba(14, 165, 233, 0.32) 0%, #0b0f17 85%) !important;
            color: #bae6fd !important;
        }
        .board-cell.explored.room-type-puzzle {
            border-color: #a855f7 !important;
            background: radial-gradient(circle at 50% 50%, rgba(168, 85, 247, 0.28) 0%, #0b0f17 85%) !important;
            color: #e9d5ff !important;
        }
        /* Cutscene Visual Stage */
        #cutscene-visual-stage {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 14px;
            margin-bottom: 18px;
        }
        .cutscene-tower-svg {
            width: clamp(84px, 18vw, 120px);
            height: clamp(84px, 18vw, 120px);
            filter: drop-shadow(0 0 18px rgba(220, 38, 38, 0.75));
        }
        .win-theme .cutscene-tower-svg {
            filter: drop-shadow(0 0 18px rgba(251, 191, 36, 0.8));
        }
        .cutscene-avatars-row {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 12px;
            flex-wrap: wrap;
        }
        .cutscene-avatar-item {
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 4px;
            background: rgba(9, 13, 20, 0.75);
            padding: 8px 12px;
            border-radius: 10px;
            border: 1px solid rgba(255, 255, 255, 0.2);
        }
        .cutscene-avatar-item img {
            margin: 0 !important;
            width: 52px;
            height: 52px;
        }
        .cutscene-avatar-name {
            font-size: 0.85rem;
            color: #f1f5f9;
            font-weight: 700;
            max-width: 110px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }

        /* ==================== Real-Event Visual Effects (Decorative, Non-blocking, aria-hidden) ==================== */
        #et-vfx-layer {
            position: fixed;
            inset: 0;
            pointer-events: none !important;
            user-select: none !important;
            z-index: 950;
            overflow: hidden;
        }
        .et-vfx-flash {
            position: absolute;
            inset: 0;
            pointer-events: none !important;
            animation: etVfxFlashFade 0.8s ease-out forwards;
        }
        .et-vfx-flash.danger {
            background: radial-gradient(circle at 50% 45%, rgba(239, 68, 68, 0.28) 0%, rgba(185, 28, 28, 0.48) 100%);
        }
        .et-vfx-flash.laser {
            background: radial-gradient(circle at 50% 45%, rgba(56, 189, 248, 0.35) 0%, rgba(3, 105, 161, 0.22) 100%);
        }
        .et-vfx-flash.shield {
            background: radial-gradient(circle at 50% 45%, rgba(96, 165, 250, 0.32) 0%, rgba(251, 191, 36, 0.2) 100%);
        }
        .et-vfx-flash.heal {
            background: radial-gradient(circle at 50% 45%, rgba(34, 197, 94, 0.26) 0%, transparent 80%);
        }
        .et-vfx-flash.gold {
            background: radial-gradient(circle at 50% 45%, rgba(251, 191, 36, 0.3) 0%, transparent 82%);
        }
        .et-vfx-flash.toilet-fog {
            background: radial-gradient(circle at 50% 55%, rgba(132, 204, 22, 0.26) 0%, rgba(107, 33, 168, 0.26) 65%, transparent 95%);
        }
        .et-vfx-flash.stairs {
            background: linear-gradient(180deg, rgba(251, 191, 36, 0.34) 0%, rgba(15, 23, 42, 0.55) 100%);
        }

        .et-vfx-stack {
            position: absolute;
            top: 28%;
            left: 50%;
            transform: translate(-50%, -50%);
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 8px;
            width: max-content;
            max-width: 92vw;
            pointer-events: none !important;
        }
        .et-vfx-card {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
            padding: 8px 18px;
            border-radius: 999px;
            background: rgba(9, 13, 20, 0.92);
            border: 2px solid var(--secondary);
            color: #ffffff;
            font-weight: 800;
            font-size: clamp(0.95rem, 3vw, 1.25rem);
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.85), 0 0 18px rgba(251, 191, 36, 0.35);
            text-shadow: 0 2px 4px rgba(0, 0, 0, 0.95);
            animation: etVfxPopFloat 1.25s cubic-bezier(0.22, 1, 0.36, 1) forwards;
            pointer-events: none !important;
            white-space: nowrap;
        }
        .et-vfx-card.danger {
            border-color: #ef4444;
            color: #fecaca;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.9), 0 0 22px rgba(239, 68, 68, 0.65);
        }
        .et-vfx-card.laser {
            border-color: #38bdf8;
            color: #e0f2fe;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.9), 0 0 22px rgba(56, 189, 248, 0.7);
        }
        .et-vfx-card.shield {
            border-color: #60a5fa;
            color: #dbeafe;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.9), 0 0 22px rgba(96, 165, 250, 0.65);
        }
        .et-vfx-card.heal {
            border-color: #22c55e;
            color: #bbf7d0;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.9), 0 0 20px rgba(34, 197, 94, 0.6);
        }
        .et-vfx-card.gold {
            border-color: #fbbf24;
            color: #fef08a;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.9), 0 0 22px rgba(251, 191, 36, 0.65);
        }
        .et-vfx-card.dirty {
            border-color: #a855f7;
            color: #e9d5ff;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.9), 0 0 20px rgba(168, 85, 247, 0.55);
        }
        .et-vfx-emoji {
            font-size: 1.35em;
            line-height: 1;
            filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.8));
        }

        /* กรงเล็บซอมบี้, ลำแสงปืนเลเซอร์, โล่ป้องกัน, ทางลงชั้น */
        .et-vfx-svg-graphic {
            width: clamp(95px, 24vw, 135px);
            height: clamp(95px, 24vw, 135px);
            pointer-events: none !important;
            filter: drop-shadow(0 4px 12px rgba(0, 0, 0, 0.85));
            animation: etVfxGraphicScale 1.05s ease-out forwards;
        }
        .et-vfx-laser-beam {
            position: absolute;
            top: 42%;
            left: 50%;
            width: min(85vw, 480px);
            height: 8px;
            transform: translate(-50%, -50%) rotate(-12deg);
            background: linear-gradient(90deg, transparent 0%, #38bdf8 20%, #ffffff 50%, #f43f5e 85%, transparent 100%);
            border-radius: 999px;
            box-shadow: 0 0 18px #38bdf8, 0 0 32px #ef4444;
            animation: etVfxLaserStrike 0.65s ease-out forwards;
            pointer-events: none !important;
        }

        /* กรอบเรืองแสงและการเน้นแถบเทิร์นและสถานะเมื่อเปลี่ยนเทิร์นหรือเกิดเหตุการณ์ */
        #game-status-bar.et-turn-active-glow {
            border-color: #fbbf24 !important;
            box-shadow: 0 0 18px rgba(251, 191, 36, 0.55), inset 0 0 10px rgba(251, 191, 36, 0.2) !important;
            animation: etTurnBannerPulse 0.85s ease-out;
        }
        #my-player-status.et-my-turn-glow {
            border-color: #fbbf24 !important;
            border-left: 5px solid #fde047 !important;
            box-shadow: 0 0 16px rgba(251, 191, 36, 0.45), inset 0 0 10px rgba(251, 191, 36, 0.14) !important;
        }
        #my-player-status.et-status-hit-shake {
            animation: etHudDamageShake 0.55s ease-in-out;
            border-color: #ef4444 !important;
        }
        #my-player-status.et-status-heal-glow {
            animation: etHudHealPulse 0.75s ease-out;
            border-color: #22c55e !important;
        }
        #latest-event-display.et-event-pulse-danger {
            border-left-color: #ef4444 !important;
            box-shadow: 0 0 14px rgba(239, 68, 68, 0.4), inset 0 1px 4px rgba(0, 0, 0, 0.7) !important;
        }
        #latest-event-display.et-event-pulse-good {
            border-left-color: #22c55e !important;
            box-shadow: 0 0 14px rgba(34, 197, 94, 0.35), inset 0 1px 4px rgba(0, 0, 0, 0.7) !important;
        }
        #latest-event-display.et-event-pulse-gold {
            border-left-color: #fbbf24 !important;
            box-shadow: 0 0 14px rgba(251, 191, 36, 0.4), inset 0 1px 4px rgba(0, 0, 0, 0.7) !important;
        }

        /* เอฟเฟกต์เน้นห้องที่เพิ่งถูกสำรวจบนกระดาน */
        .board-cell.et-room-just-explored {
            animation: etRoomPopEffect 0.85s cubic-bezier(0.22, 1, 0.36, 1);
            z-index: 3;
        }

        @keyframes etVfxFlashFade {
            0% { opacity: 0; }
            20% { opacity: 1; }
            100% { opacity: 0; }
        }
        @keyframes etVfxPopFloat {
            0% { opacity: 0; transform: translateY(16px) scale(0.82); }
            22% { opacity: 1; transform: translateY(0) scale(1.05); }
            78% { opacity: 1; transform: translateY(-6px) scale(1); }
            100% { opacity: 0; transform: translateY(-18px) scale(0.94); }
        }
        @keyframes etVfxGraphicScale {
            0% { opacity: 0; transform: scale(0.65); }
            25% { opacity: 1; transform: scale(1.08); }
            75% { opacity: 1; transform: scale(1); }
            100% { opacity: 0; transform: scale(1.12); }
        }
        @keyframes etVfxLaserStrike {
            0% { opacity: 0; transform: translate(-50%, -50%) rotate(-12deg) scaleX(0.1); }
            30% { opacity: 1; transform: translate(-50%, -50%) rotate(-12deg) scaleX(1.05); }
            100% { opacity: 0; transform: translate(-50%, -50%) rotate(-12deg) scaleX(1); }
        }
        @keyframes etTurnBannerPulse {
            0% { transform: scale(0.97); }
            50% { transform: scale(1.02); }
            100% { transform: scale(1); }
        }
        @keyframes etHudDamageShake {
            0%, 100% { transform: translateX(0); }
            20%, 60% { transform: translateX(-5px); }
            40%, 80% { transform: translateX(5px); }
        }
        @keyframes etHudHealPulse {
            0% { box-shadow: 0 0 0 rgba(34, 197, 94, 0); }
            50% { box-shadow: 0 0 18px rgba(34, 197, 94, 0.65); }
            100% { box-shadow: 0 0 0 rgba(34, 197, 94, 0); }
        }
        @keyframes etRoomPopEffect {
            0% { transform: scale(0.9); box-shadow: 0 0 0 rgba(251, 191, 36, 0); }
            45% { transform: scale(1.08); box-shadow: 0 0 22px rgba(251, 191, 36, 0.75); }
            100% { transform: scale(1); }
        }
    `;
    document.head.appendChild(style);
}

// ==================== Decorative Visual Effects Engine (เชื่อมกับเหตุการณ์จริงเท่านั้น & ซ่อนจาก SC 100%) ====================
function getOrCreateVfxLayer() {
    ensureVisualRuntimeStyles();
    let layer = document.getElementById('et-vfx-layer');
    if (!layer) {
        layer = document.createElement('div');
        layer.id = 'et-vfx-layer';
        layer.setAttribute('aria-hidden', 'true');
        document.body.appendChild(layer);
    }
    return layer;
}

function getClawSvgHtml() {
    return `<svg class="et-vfx-svg-graphic" viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg" focusable="false" aria-hidden="true">
        <path d="M18 22 L98 92 M38 14 L106 76 M12 44 L84 106" stroke="#ef4444" stroke-width="7" stroke-linecap="round"/>
        <path d="M18 22 L98 92 M38 14 L106 76 M12 44 L84 106" stroke="#fecaca" stroke-width="2.5" stroke-linecap="round"/>
    </svg>`;
}

function getShieldSvgHtml() {
    return `<svg class="et-vfx-svg-graphic" viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg" focusable="false" aria-hidden="true">
        <path d="M60 12 L100 28 V58 C100 84 82 102 60 110 C38 102 20 84 20 58 V28 L60 12 Z" fill="rgba(37, 99, 235, 0.38)" stroke="#60a5fa" stroke-width="5"/>
        <path d="M60 22 L90 34 V56 C90 76 76 90 60 97 C44 90 30 76 30 56 V34 L60 22 Z" fill="none" stroke="#fde047" stroke-width="2.5"/>
    </svg>`;
}

function getLaserSvgHtml() {
    return `<svg class="et-vfx-svg-graphic" viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg" focusable="false" aria-hidden="true">
        <circle cx="60" cy="60" r="38" fill="none" stroke="#38bdf8" stroke-width="4" stroke-dasharray="10 6"/>
        <circle cx="60" cy="60" r="16" fill="rgba(239, 68, 68, 0.65)" stroke="#ffffff" stroke-width="3"/>
        <line x1="60" y1="8" x2="60" y2="112" stroke="#38bdf8" stroke-width="3"/>
        <line x1="8" y1="60" x2="112" y2="60" stroke="#38bdf8" stroke-width="3"/>
    </svg>`;
}

function getDoorPortalSvgHtml() {
    return `<svg class="et-vfx-svg-graphic" viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg" focusable="false" aria-hidden="true">
        <rect x="28" y="14" width="64" height="92" rx="4" fill="#0f172a" stroke="#fbbf24" stroke-width="4"/>
        <path d="M32 18 L74 26 V94 L32 102 Z" fill="rgba(251, 191, 36, 0.78)"/>
        <circle cx="66" cy="60" r="4" fill="#0f172a"/>
    </svg>`;
}

function spawnVfxPresentation({ flashType = null, cards = [], svgHtml = '', showLaserBeam = false }) {
    const layer = getOrCreateVfxLayer();
    layer.innerHTML = '';

    if (flashType) {
        const flashEl = document.createElement('div');
        flashEl.className = `et-vfx-flash ${flashType}`;
        flashEl.setAttribute('aria-hidden', 'true');
        layer.appendChild(flashEl);
    }

    if (showLaserBeam) {
        const beamEl = document.createElement('div');
        beamEl.className = 'et-vfx-laser-beam';
        beamEl.setAttribute('aria-hidden', 'true');
        layer.appendChild(beamEl);
    }

    const stack = document.createElement('div');
    stack.className = 'et-vfx-stack';
    stack.setAttribute('aria-hidden', 'true');

    if (svgHtml) {
        const svgWrap = document.createElement('div');
        svgWrap.setAttribute('aria-hidden', 'true');
        svgWrap.innerHTML = svgHtml;
        stack.appendChild(svgWrap);
    }

    cards.forEach((c) => {
        const pill = document.createElement('div');
        pill.className = `et-vfx-card ${c.theme || ''}`;
        pill.setAttribute('aria-hidden', 'true');
        pill.innerHTML = `<span class="et-vfx-emoji" aria-hidden="true">${c.icon || '✨'}</span><span aria-hidden="true">${escapeHtml(c.text)}</span>`;
        stack.appendChild(pill);
    });

    layer.appendChild(stack);

    setTimeout(() => {
        if (layer.contains(stack)) {
            layer.innerHTML = '';
        }
    }, 1350);
}

function pulseEventBoxTheme(theme) {
    const evBox = document.getElementById('latest-event-display');
    if (!evBox) return;
    evBox.classList.remove('et-event-pulse-danger', 'et-event-pulse-good', 'et-event-pulse-gold');
    if (theme) {
        void evBox.offsetWidth;
        evBox.classList.add(theme);
    }
}

function pulseMyStatusHud(mode) {
    const myHud = document.getElementById('my-player-status');
    if (!myHud) return;
    myHud.classList.remove('et-status-hit-shake', 'et-status-heal-glow');
    if (mode) {
        void myHud.offsetWidth;
        myHud.classList.add(mode);
        setTimeout(() => {
            myHud.classList.remove(mode);
        }, 800);
    }
}

function highlightExploredRoomCell(roomNum1Based) {
    setTimeout(() => {
        const boardEl = document.getElementById('escape-board');
        if (!boardEl) return;
        const cells = boardEl.querySelectorAll('.board-cell');
        const target = cells[roomNum1Based - 1];
        if (target) {
            target.classList.remove('et-room-just-explored');
            void target.offsetWidth;
            target.classList.add('et-room-just-explored');
        }
    }, 50);
}

function triggerTurnChangeVisual(turnKey, activePlayerId, activePlayerName) {
    if (!turnKey || lastVfxTurnKey === turnKey) return;
    lastVfxTurnKey = turnKey;

    const statusBar = document.getElementById('game-status-bar');
    if (statusBar) {
        statusBar.classList.remove('et-turn-active-glow');
        void statusBar.offsetWidth;
        statusBar.classList.add('et-turn-active-glow');
    }

    const myHud = document.getElementById('my-player-status');
    if (myHud) {
        if (activePlayerId === myId) {
            myHud.classList.add('et-my-turn-glow');
        } else {
            myHud.classList.remove('et-my-turn-glow');
        }
    }

    if (activePlayerName) {
        spawnVfxPresentation({
            flashType: activePlayerId === myId ? 'gold' : null,
            cards: [{
                icon: activePlayerId === myId ? '⚡' : '⏳',
                text: `เทิร์นของ ${activePlayerName}`,
                theme: 'gold'
            }]
        });
    }
}

function triggerVisualEffectFromLog(logObj) {
    if (!logObj || !logObj.ts || logObj.ts <= lastVfxLogTs) return;
    lastVfxLogTs = logObj.ts;

    const msg = String(logObj.msg || '').trim();
    if (!msg) return;

    const myPlayerName = (gameState && gameState.players && myId && gameState.players[myId])
        ? gameState.players[myId].name
        : '';
    const isMeInvolved = myPlayerName && msg.startsWith(myPlayerName);

    // ตรวจสอบว่ามีการสำรวจห้องหมายเลขใด เพื่อแสดงเอฟเฟกต์เน้นที่ปุ่มห้องนั้นบนกระดาน (เฉพาะชั้นเดียวกัน)
    const roomMatch = msg.match(/สำรวจห้อง\s+(\d+):/);
    if (roomMatch) {
        const roomNum = parseInt(roomMatch[1], 10);
        if (roomNum >= 1 && roomNum <= 20) {
            highlightExploredRoomCell(roomNum);
        }
    }

    const cards = [];
    let flashType = null;
    let svgHtml = '';
    let showLaserBeam = false;
    let eventBoxTheme = null;

    // 1. เหตุการณ์ช่วงเริ่มเทิร์น (ใช้วัคซีนอัตโนมัติ / กินอาหารจากคลัง / เสีย HP จากการติดเชื้อ 20 หรือ 40 / พักฟื้น)
    if (msg.includes('ใช้วัคซีนรักษาการติดเชื้ออัตโนมัติ')) {
        cards.push({ icon: '💉', text: 'ใช้วัคซีน รักษาการติดเชื้อ', theme: 'heal' });
        flashType = 'heal';
        eventBoxTheme = 'et-event-pulse-good';
        if (isMeInvolved) pulseMyStatusHud('et-status-heal-glow');
    }

    const startFoodMatch = msg.match(/กินอาหารฟื้นฟู HP (\d+) หน่วย/);
    if (startFoodMatch && !msg.includes('สำรวจห้อง')) {
        const healAmt = startFoodMatch[1];
        cards.push({ icon: '🥫', text: `+${healAmt} HP (กินอาหาร)`, theme: 'heal' });
        flashType = 'heal';
        eventBoxTheme = 'et-event-pulse-good';
        if (isMeInvolved) pulseMyStatusHud('et-status-heal-glow');
    }

    const infDmgMatch = msg.match(/เสีย (\d+) HP จากการติดเชื้อ/);
    if (infDmgMatch) {
        const dmgAmt = infDmgMatch[1];
        cards.push({ icon: '☣️', text: `-${dmgAmt} HP (พิษติดเชื้อ)`, theme: 'danger' });
        svgHtml = getClawSvgHtml();
        flashType = 'danger';
        eventBoxTheme = 'et-event-pulse-danger';
        if (isMeInvolved) pulseMyStatusHud('et-status-hit-shake');
    }

    if (msg.includes('พักฟื้นเสร็จสิ้น ฟื้นฟู HP เป็น 150')) {
        cards.push({ icon: '💚', text: 'ฟื้นฟู HP เป็น 150', theme: 'heal' });
        flashType = 'heal';
        eventBoxTheme = 'et-event-pulse-good';
        if (isMeInvolved) pulseMyStatusHud('et-status-heal-glow');
    } else if (msg.includes('กำลังพักฟื้น หยุดเล่น 1 เทิร์น') || msg.includes('ได้รับบาดเจ็บสาหัส ต้องหยุดพักฟื้น')) {
        cards.push({ icon: '🛌', text: 'บาดเจ็บสาหัส พักฟื้น 1 เทิร์น', theme: 'danger' });
        flashType = flashType || 'danger';
        eventBoxTheme = 'et-event-pulse-danger';
    }

    // 2. เหตุการณ์เจอซอมบี้: ยิงซอมบี้ / เกราะป้องกัน / ถูกโจมตีติดเชื้อ
    if (msg.includes('พบซอมบี้!') || msg.includes('เจอซอมบี้!')) {
        if (msg.includes('ใช้ปืนยิงป้องกันสำเร็จ') || msg.includes('ยิงป้องกันสำเร็จ')) {
            showLaserBeam = true;
            svgHtml = getLaserSvgHtml();
            flashType = 'laser';
            eventBoxTheme = 'et-event-pulse-gold';
            cards.push({ icon: '🔫', text: 'ยิงเลเซอร์สกัดซอมบี้! (-1 กระสุน)', theme: 'laser' });
        } else if (msg.includes('เกราะสวรรค์ป้องกันการโจมตีสำเร็จ') || msg.includes('เกราะกันสำเร็จ')) {
            svgHtml = getShieldSvgHtml();
            flashType = 'shield';
            eventBoxTheme = 'et-event-pulse-gold';
            cards.push({ icon: '🛡️', text: 'เกราะป้องกันซอมบี้สำเร็จ! (-1 เกราะ)', theme: 'shield' });
        } else if (msg.includes('ถูกโจมตี ติดเชื้อระดับ') || msg.includes('ติดเชื้อ!')) {
            svgHtml = getClawSvgHtml();
            flashType = 'danger';
            eventBoxTheme = 'et-event-pulse-danger';
            const infLevelMatch = msg.match(/ติดเชื้อระดับ (\d+) HP\/เทิร์น/);
            const infText = infLevelMatch
                ? `ถูกซอมบี้ข่วน! ติดเชื้อ -${infLevelMatch[1]} HP/เทิร์น`
                : 'ถูกซอมบี้ข่วน! ติดเชื้อ!';
            cards.push({ icon: '🧟', text: infText, theme: 'danger' });
            if (isMeInvolved) pulseMyStatusHud('et-status-hit-shake');
        }
    }

    // 3. เหตุการณ์พบสิ่งของหรือห้องพิเศษจากการสำรวจห้องจริง
    if (msg.includes('พบกุญแจปกติ')) {
        flashType = 'gold';
        eventBoxTheme = 'et-event-pulse-gold';
        cards.push({ icon: '🔑', text: '+1 กุญแจปกติ', theme: 'gold' });
    } else if (msg.includes('พบกุญแจผี')) {
        flashType = 'gold';
        eventBoxTheme = 'et-event-pulse-gold';
        cards.push({ icon: '🗝️', text: '+1 กุญแจผี', theme: 'gold' });
    } else if (msg.includes('พบทางลงลับ') || msg.includes('พบทางลับ')) {
        flashType = 'gold';
        svgHtml = getDoorPortalSvgHtml();
        eventBoxTheme = 'et-event-pulse-gold';
        cards.push({ icon: '🪜', text: 'พบทางลงลับ! (ลงข้ามชั้นได้)', theme: 'gold' });
    } else if (msg.includes('พบวัคซีน')) {
        flashType = 'heal';
        eventBoxTheme = 'et-event-pulse-good';
        cards.push({ icon: '💉', text: '+1 วัคซีน', theme: 'heal' });
    } else if (msg.includes('พบอาหาร') || msg.includes('กินอาหารฟื้น HP') || msg.includes('เก็บอาหารเข้าคลังสำรอง')) {
        flashType = 'heal';
        eventBoxTheme = 'et-event-pulse-good';
        const instantHealMatch = msg.match(/กินทันทีฟื้นฟู HP (\d+) หน่วย/);
        if (instantHealMatch) {
            cards.push({ icon: '🥫', text: `พบอาหาร +${instantHealMatch[1]} HP`, theme: 'heal' });
            if (isMeInvolved) pulseMyStatusHud('et-status-heal-glow');
        } else if (msg.includes('กินอาหารฟื้น HP')) {
            cards.push({ icon: '🥫', text: 'พบอาหาร กินฟื้นฟู HP', theme: 'heal' });
        } else {
            cards.push({ icon: '🥫', text: '+1 อาหารเก็บเข้าคลัง', theme: 'heal' });
        }
    } else if (msg.includes('พบเกราะสวรรค์') || msg.includes('ได้เกราะ')) {
        flashType = 'shield';
        svgHtml = getShieldSvgHtml();
        eventBoxTheme = 'et-event-pulse-gold';
        cards.push({ icon: '🛡️', text: '+1 เกราะสวรรค์', theme: 'shield' });
    } else if (msg.includes('พบห้องน้ำเก่าสกปรกแถมผุพัง')) {
        flashType = 'toilet-fog';
        eventBoxTheme = 'et-event-pulse-danger';
        cards.push({ icon: '🚽', text: 'ห้องน้ำเก่าและสกปรก!', theme: 'dirty' });
    } else if (msg.includes('พบปืนเลเซอร์') || msg.includes('ได้ปืนและกระสุน')) {
        flashType = 'laser';
        eventBoxTheme = 'et-event-pulse-gold';
        if (msg.includes('แต่กระสุนเต็มแล้ว')) {
            cards.push({ icon: '🔫', text: 'พบปืนเลเซอร์ (กระสุนเต็มแล้ว)', theme: 'laser' });
        } else if (msg.includes('แปลงเป็นกระสุน 1 นัด')) {
            cards.push({ icon: '🔋', text: '+1 กระสุนเลเซอร์', theme: 'laser' });
        } else {
            cards.push({ icon: '🔫', text: 'ได้ปืนเลเซอร์ + กระสุน!', theme: 'laser' });
        }
    } else if (msg.includes('พบแบตเตอรี่') || msg.includes('ได้กระสุน')) {
        flashType = 'laser';
        eventBoxTheme = 'et-event-pulse-gold';
        if (msg.includes('แต่กระสุนเต็มแล้ว')) {
            cards.push({ icon: '🔋', text: 'พบแบตเตอรี่ (กระสุนเต็มแล้ว)', theme: 'laser' });
        } else {
            cards.push({ icon: '🔋', text: '+1 กระสุนเลเซอร์', theme: 'laser' });
        }
    }

    // 4. เหตุการณ์ตอบคำใบ้ถูกต้อง
    if (msg.includes('ตอบคำใบ้ถูกต้อง')) {
        flashType = 'gold';
        eventBoxTheme = 'et-event-pulse-gold';
        cards.push({ icon: '🔓', text: 'ถอดรหัสคำใบ้สำเร็จ!', theme: 'gold' });
    }

    // 5. เหตุการณ์เปิดประตูลงชั้นถัดไป หรือ เปิดประตูหนีออกจากตึกชั้น 1
    const stairsMatch = msg.match(/ใช้(.+?)\s+ลงมายังชั้น\s+(\d+)/);
    if (stairsMatch) {
        const usedItem = stairsMatch[1];
        const targetFloor = stairsMatch[2];
        flashType = 'stairs';
        svgHtml = getDoorPortalSvgHtml();
        eventBoxTheme = 'et-event-pulse-gold';
        cards.push({ icon: '🚪', text: `เปิดประตูลงสู่ชั้น ${targetFloor} (${usedItem})`, theme: 'gold' });
    } else if (msg.includes('ปลดล็อกประตูทางออกชั้น 1')) {
        flashType = 'gold';
        svgHtml = getDoorPortalSvgHtml();
        eventBoxTheme = 'et-event-pulse-gold';
        cards.push({ icon: '🏆', text: 'เปิดประตูทางออกตึกสำเร็จ!', theme: 'gold' });
    }

    if (eventBoxTheme) {
        pulseEventBoxTheme(eventBoxTheme);
    }

    if (cards.length > 0 || svgHtml || showLaserBeam) {
        spawnVfxPresentation({
            flashType,
            cards,
            svgHtml,
            showLaserBeam
        });
    }
}

// Audio System
const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
const unlockAudio = () => {
    if (audioCtx.state === 'suspended' || audioCtx.state === 'interrupted') {
        audioCtx.resume().catch(() => {});
    }
};
document.addEventListener('click', unlockAudio, { capture: true });
document.addEventListener('keydown', unlockAudio, { capture: true });
document.addEventListener('touchstart', unlockAudio, { capture: true, passive: true });
document.addEventListener('touchend', unlockAudio, { capture: true, passive: true });

let audioPlaybackQueue = Promise.resolve();

function playAudio(name) {
    return new Promise((resolve) => {
        try {
            if (audioCtx.state === 'suspended' || audioCtx.state === 'interrupted') {
                audioCtx.resume().catch(() => {});
            }
            fetch(`audio/${name}`)
                .then(res => res.arrayBuffer())
                .then(buf => audioCtx.decodeAudioData(buf))
                .then(decoded => {
                    const source = audioCtx.createBufferSource();
                    source.buffer = decoded;
                    source.connect(audioCtx.destination);
                    source.onended = () => resolve();
                    source.start(0);
                }).catch(() => resolve());
        } catch(e) {
            resolve();
        }
    });
}

function playAudioSequence(files) {
    if (!files) return Promise.resolve();
    const list = Array.isArray(files) ? files : [files];
    audioPlaybackQueue = audioPlaybackQueue.then(async () => {
        for (const f of list) {
            if (f) await playAudio(f);
        }
    }).catch(() => {});
    return audioPlaybackQueue;
}

function playBGM() {
    stopBGM();
    isBgmPlaying = true;
    try {
        if (audioCtx.state === 'suspended' || audioCtx.state === 'interrupted') {
            audioCtx.resume().catch(() => {});
        }
        fetch('audio/bgm.mp3')
            .then(res => res.arrayBuffer())
            .then(buf => audioCtx.decodeAudioData(buf))
            .then(decoded => {
                if (!isBgmPlaying) return;
                bgmSource = audioCtx.createBufferSource();
                bgmSource.buffer = decoded;
                bgmSource.loop = true;
                bgmSource.connect(audioCtx.destination);
                bgmSource.start(0);
            }).catch(() => {});
    } catch(e){}
}

function stopBGM() {
    isBgmPlaying = false;
    if (bgmSource) {
        try {
            bgmSource.stop(0);
            bgmSource.disconnect();
        } catch(e) {}
        bgmSource = null;
    }
}

window.addEventListener('beforeunload', () => {
    stopBGM();
});

document.addEventListener('visibilitychange', () => {
    if (document.hidden) {
        if (audioCtx.state === 'running') {
            audioCtx.suspend().catch(() => {});
        }
    } else {
        if (audioCtx.state === 'suspended' || audioCtx.state === 'interrupted') {
            audioCtx.resume().catch(() => {});
        }
    }
});

window.addEventListener('pageshow', () => {
    if (!document.hidden && (audioCtx.state === 'suspended' || audioCtx.state === 'interrupted')) {
        audioCtx.resume().catch(() => {});
    }
});

window.addEventListener('focus', () => {
    if (!document.hidden && (audioCtx.state === 'suspended' || audioCtx.state === 'interrupted')) {
        audioCtx.resume().catch(() => {});
    }
});

// Scene Animation & Synchronized ARIA Live Helper
function clearSceneTimeouts() {
    sceneTimeouts.forEach(t => clearTimeout(t));
    sceneTimeouts = [];
    const overlay = document.getElementById('cutscene-overlay');
    if (overlay) overlay.remove();
}

function buildCutsceneVisualStageHtml(theme) {
    if (!gameState || !gameState.players) return '';
    const orderKeys = (gameState.order && gameState.order.length > 0)
        ? gameState.order
        : Object.keys(gameState.players);

    if (theme === 'scary') {
        let avatarsHtml = '';
        orderKeys.forEach(pid => {
            const p = gameState.players[pid];
            if (!p) return;
            avatarsHtml += `
                <div class="cutscene-avatar-item" aria-hidden="true">
                    ${getParticipantAvatarHtml(pid, p)}
                    <span class="cutscene-avatar-name">${escapeHtml(p.name)}</span>
                </div>
            `;
        });
        return `
            <svg class="cutscene-tower-svg" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg" focusable="false" aria-hidden="true">
                <rect x="24" y="10" width="52" height="80" rx="4" fill="#131b26" stroke="#dc2626" stroke-width="3"/>
                <line x1="24" y1="23" x2="76" y2="23" stroke="#2e3c51" stroke-width="2"/>
                <line x1="24" y1="36" x2="76" y2="36" stroke="#2e3c51" stroke-width="2"/>
                <line x1="24" y1="49" x2="76" y2="49" stroke="#2e3c51" stroke-width="2"/>
                <line x1="24" y1="62" x2="76" y2="62" stroke="#2e3c51" stroke-width="2"/>
                <line x1="24" y1="75" x2="76" y2="75" stroke="#2e3c51" stroke-width="2"/>
                <rect x="34" y="14" width="8" height="5" fill="#fbbf24"/>
                <rect x="58" y="27" width="8" height="5" fill="#dc2626"/>
                <rect x="34" y="40" width="8" height="5" fill="#fbbf24"/>
                <rect x="44" y="78" width="12" height="12" fill="#fbbf24"/>
            </svg>
            <div class="cutscene-avatars-row" aria-hidden="true">${avatarsHtml}</div>
        `;
    } else {
        const winId = gameState.winner;
        const winP = winId && gameState.players[winId] ? gameState.players[winId] : null;
        let winAvatarHtml = '';
        if (winP) {
            winAvatarHtml = `
                <div class="cutscene-avatar-item" aria-hidden="true">
                    ${getParticipantAvatarHtml(winId, winP)}
                    <span class="cutscene-avatar-name">🏆 ${escapeHtml(winP.name)}</span>
                </div>
            `;
        }
        return `
            <svg class="cutscene-tower-svg" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg" focusable="false" aria-hidden="true">
                <circle cx="50" cy="50" r="42" fill="#131b26" stroke="#fbbf24" stroke-width="3"/>
                <rect x="34" y="26" width="32" height="48" rx="3" fill="#0f172a" stroke="#22c55e" stroke-width="2.5"/>
                <path d="M34 26 L56 20 L56 68 L34 74 Z" fill="#fbbf24" opacity="0.85"/>
            </svg>
            <div class="cutscene-avatars-row" aria-hidden="true">${winAvatarHtml}</div>
        `;
    }
}

function createCutsceneOverlay(theme) {
    ensureVisualRuntimeStyles();
    let existing = document.getElementById('cutscene-overlay');
    if (existing) existing.remove();

    if (!document.getElementById('cutscene-style')) {
        const style = document.createElement('style');
        style.id = 'cutscene-style';
        style.textContent = `
            @keyframes pulseSceneText {
                0% { transform: scale(0.85); opacity: 0; }
                20% { transform: scale(1.05); opacity: 1; }
                85% { transform: scale(1); opacity: 1; }
                100% { transform: scale(0.95); opacity: 0.9; }
            }
            @keyframes scaryBgPulse {
                0%, 100% { background: radial-gradient(circle, #3a0000 0%, #050000 85%); }
                50% { background: radial-gradient(circle, #590000 0%, #0a0000 85%); }
            }
            @keyframes winBgGlow {
                0%, 100% { background: radial-gradient(circle, #1b4d20 0%, #041206 85%); }
                50% { background: radial-gradient(circle, #2e7d32 0%, #08240c 85%); }
            }
            #cutscene-overlay {
                position: fixed;
                inset: 0;
                z-index: 9999;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                padding: 20px;
                text-align: center;
                user-select: none;
            }
            #cutscene-overlay.scary-theme {
                animation: scaryBgPulse 2.5s infinite ease-in-out;
                box-shadow: inset 0 0 120px #000;
            }
            #cutscene-overlay.win-theme {
                animation: winBgGlow 2.5s infinite ease-in-out;
                box-shadow: inset 0 0 100px #000;
            }
            #cutscene-text {
                font-size: clamp(2.2rem, 8vw, 4.5rem);
                font-weight: 900;
                max-width: 92vw;
                word-wrap: break-word;
                line-height: 1.35;
                letter-spacing: 2px;
            }
            .scary-theme #cutscene-text {
                color: #ffebee;
                text-shadow: 0 0 15px #ff1744, 0 0 35px #b71c1c, 2px 2px 5px #000;
            }
            .win-theme #cutscene-text {
                color: #fff8e1;
                text-shadow: 0 0 15px #ffb300, 0 0 35px #2e7d32, 2px 2px 5px #000;
            }
        `;
        document.head.appendChild(style);
    }

    const overlay = document.createElement('div');
    overlay.id = 'cutscene-overlay';
    overlay.className = theme === 'scary' ? 'scary-theme' : 'win-theme';
    overlay.setAttribute('aria-hidden', 'true');

    const visualStage = document.createElement('div');
    visualStage.id = 'cutscene-visual-stage';
    visualStage.setAttribute('aria-hidden', 'true');
    visualStage.innerHTML = buildCutsceneVisualStageHtml(theme);
    overlay.appendChild(visualStage);

    const textEl = document.createElement('div');
    textEl.id = 'cutscene-text';
    overlay.appendChild(textEl);

    document.body.appendChild(overlay);
    return textEl;
}

function announceSceneCue(text, clearDelay = 850) {
    speechQueue = [];
    isSpeaking = false;
    const assertiveEl = document.getElementById('sr-assertive');
    const politeEl = document.getElementById('sr-polite');
    if (politeEl) politeEl.textContent = '';
    if (assertiveEl) {
        assertiveEl.textContent = '';
        const t1 = setTimeout(() => {
            assertiveEl.textContent = text;
            const t2 = setTimeout(() => {
                if (assertiveEl.textContent === text) {
                    assertiveEl.textContent = '';
                }
            }, clearDelay);
            sceneTimeouts.push(t2);
        }, 40);
        sceneTimeouts.push(t1);
    }
}

function runStartCutscene(onComplete) {
    clearSceneTimeouts();
    stopBGM();
    playAudio('start.mp3');

    const textEl = createCutsceneOverlay('scary');
    const updateText = (msg) => {
        textEl.style.animation = 'none';
        void textEl.offsetWidth;
        textEl.textContent = msg;
        textEl.style.animation = 'pulseSceneText 1.2s ease-out forwards';
    };

    // 1.0 วินาที = ตั้งสติ
    sceneTimeouts.push(setTimeout(() => {
        updateText('ตั้งสติ');
        announceSceneCue('ตั้งสติ', 1500);
    }, 1000));

    // 3.0 วินาที = ต้องรอด
    sceneTimeouts.push(setTimeout(() => {
        updateText('ต้องรอด');
        announceSceneCue('ต้องรอด', 950);
    }, 3000));

    // 4.2 วินาที = ไป
    sceneTimeouts.push(setTimeout(() => {
        updateText('ไป');
        announceSceneCue('ไป', 750);
    }, 4200));

    // 5.10 วินาที = จบซีน เข้าสู่หน้าเกม และรอ 0.2 วินาทีเพื่อเล่น bgm.mp3
    sceneTimeouts.push(setTimeout(() => {
        const overlay = document.getElementById('cutscene-overlay');
        if (overlay) overlay.remove();
        const assertiveEl = document.getElementById('sr-assertive');
        if (assertiveEl) assertiveEl.textContent = '';
        
        if (onComplete) onComplete();

        sceneTimeouts.push(setTimeout(() => {
            if (gameState && gameState.status === 'playing') {
                playBGM();
            }
        }, 200));
    }, 5100));
}

function runWinCutscene(winnerName, onComplete) {
    clearSceneTimeouts();
    stopBGM();
    playAudio('win.mp3');

    const textEl = createCutsceneOverlay('win');
    const updateText = (msg) => {
        textEl.style.animation = 'none';
        void textEl.offsetWidth;
        textEl.textContent = msg;
        textEl.style.animation = 'pulseSceneText 1.5s ease-out forwards';
    };

    // 1.0 วินาที = ยินดีด้วย
    sceneTimeouts.push(setTimeout(() => {
        updateText('ยินดีด้วย');
        announceSceneCue('ยินดีด้วย', 1600);
    }, 1000));

    // 3.0 วินาที = (ชื่อตัวละคร) รอดชีวิต
    sceneTimeouts.push(setTimeout(() => {
        const msg = `${winnerName} รอดชีวิต`;
        updateText(msg);
        announceSceneCue(msg, 1700);
    }, 3000));

    // 5.0 วินาที = จบซีน ตัดเข้าหน้าจบเกมตามปกติ
    sceneTimeouts.push(setTimeout(() => {
        const overlay = document.getElementById('cutscene-overlay');
        if (overlay) overlay.remove();
        const assertiveEl = document.getElementById('sr-assertive');
        if (assertiveEl) assertiveEl.textContent = '';
        
        if (onComplete) onComplete();
    }, 5000));
}

// ARIA Announcement (จัดการผ่าน Central Queue ป้องกันเสียงพูดชนกัน)
function announceSR(text, priority = 'polite') {
    if (!text) return;
    if (speechQueue.length > 0 && speechQueue[speechQueue.length - 1].text === text) return;
    speechQueue.push({ text, priority });
    processSpeechQueue();
}

function processSpeechQueue() {
    if (isSpeaking || speechQueue.length === 0) return;
    isSpeaking = true;
    const item = speechQueue.shift();
    const targetEl = document.getElementById(item.priority === 'assertive' ? 'sr-assertive' : 'sr-polite');
    if (targetEl) {
        targetEl.textContent = '';
        setTimeout(() => {
            targetEl.textContent = item.text;
            const duration = Math.max(1400, item.text.length * 45);
            setTimeout(() => { 
                targetEl.textContent = ''; 
                isSpeaking = false; 
                processSpeechQueue(); 
            }, duration);
        }, 70);
    } else { 
        isSpeaking = false; 
    }
}

// Screen Switcher
function switchScreen(screenId, focusId = null) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active-screen'));
    const target = document.getElementById(screenId);
    if (target) {
        target.classList.add('active-screen');
        if (focusId) {
            const el = document.getElementById(focusId);
            if (el) { el.setAttribute('tabindex', '-1'); el.focus(); }
        }
    }
}

// Initial Events
document.addEventListener('DOMContentLoaded', () => {
    ensureVisualRuntimeStyles();
    initLobbyListener();
    const nameInput = document.getElementById('player-name-input');
    const btnConfirm = document.getElementById('btn-confirm-name');

    nameInput.addEventListener('input', (e) => {
        const hasVal = e.target.value.trim().length > 0;
        btnConfirm.disabled = !hasVal;
        btnConfirm.setAttribute('aria-disabled', (!hasVal).toString());
    });
    nameInput.addEventListener('keydown', (e) => {
        if(e.key === 'Enter' && e.target.value.trim().length > 0) confirmName();
    });
    btnConfirm.onclick = confirmName;
    document.getElementById('btn-create-room').onclick = createRoom;
    document.getElementById('btn-leave-room').onclick = window.leaveRoomAndGoHome;
    document.getElementById('btn-start-game').onclick = startGame;
    
    document.getElementById('btn-open-manual-welcome').onclick = () => showManual(true);
    document.getElementById('btn-open-manual-lobby').onclick = () => showManual(true);
    document.getElementById('btn-close-manual').onclick = () => showManual(false);

    document.getElementById('btn-add-bot').onclick = () => adjustBot(1);
    document.getElementById('btn-remove-bot').onclick = () => adjustBot(-1);
    
    document.getElementById('btn-read-all-status').onclick = readAllStatus;
    document.getElementById('btn-close-all-status').onclick = closeAllStatus;
    
    document.getElementById('btn-close-clue').onclick = async () => {
        const m = document.getElementById('clue-modal');
        const inputEl = document.getElementById('clue-answer-input');
        if (m.style.display === 'none' || inputEl.disabled) return;
        
        if (!gameState) return;
        const thisTurnKey = `${gameState.turnIndex}_${myId}`;
        if (activeTurnLock === thisTurnKey) return;
        activeTurnLock = thisTurnKey;
        
        m.style.display = 'none';
        inputEl.disabled = true;
        document.getElementById('btn-submit-clue').disabled = true;
        
        await update(ref(db), { 
            [`games/Escape/rooms/${currentRoomId}/players/${myId}/answeringClue`]: null,
            [`games/Escape/rooms/${currentRoomId}/players/${myId}/actionStarted`]: Date.now()
        });
        await emitLog(`${gameState.players[myId].name} ปิดหน้าต่างคำใบ้โดยไม่ตอบคำถาม`, ['key4.mp3']);
        setTimeout(() => { 
            endTurn(myId); 
        }, 2200);
    };
    
    document.getElementById('btn-submit-clue').onclick = submitClue;
    document.getElementById('clue-answer-input').addEventListener('keydown', (e) => {
        if (e.key === 'Enter') submitClue();
    });
});

function confirmName() {
    myName = document.getElementById('player-name-input').value.trim() || 'ผู้รอดชีวิต';
    playAudio('select.mp3');
    switchScreen('screen-lobby', 'lobby-heading');
    announceSR(`เข้าสู่ล็อบบี้แล้ว ยินดีต้อนรับ ${myName}`);
}

function showManual(show) {
    const m = document.getElementById('manual-page');
    m.style.display = show ? 'flex' : 'none';
    if(show) { document.getElementById('manual-title').focus(); announceSR("เปิดคู่มือเกมแล้ว"); }
    else { announceSR("ปิดคู่มือเกมแล้ว"); }
}

// Lobby / Room Management
function initLobbyListener() {
    onValue(ref(db, 'games/Escape/rooms'), (snap) => {
        const rooms = snap.val() || {};
        const container = document.getElementById('room-list');
        if (myId) return; // Ignore if already in room
        container.innerHTML = '';
        let count = 0;
        for (const [rId, rData] of Object.entries(rooms)) {
            if (rData.status === 'waiting') {
                count++;
                const pCount = rData.players ? Object.keys(rData.players).length : 0;
                if (pCount < 4) {
                    const d = document.createElement('div');
                    d.className = 'room-item';
                    d.innerHTML = `เข้าร่วมห้อง ${rId} (มี ${pCount}/4 คน)`;
                    d.onclick = () => joinRoom(rId);
                    d.setAttribute('role', 'button');
                    d.setAttribute('tabindex', '0');
                    d.onkeydown = (e) => { if (e.key === 'Enter') joinRoom(rId); };
                    container.appendChild(d);
                }
            }
        }
        if(count===0) container.innerHTML = '<div class="loading-text">ไม่มีห้องที่รอผู้เล่น...</div>';
    });
}

function createRoom() {
    playAudio('1.mp3');
    runTransaction(ref(db, 'games/Escape/room_counter'), (cur) => (cur || 0) + 1).then((res) => {
        if (res.committed) {
            currentRoomId = 'Escape' + String(res.snapshot.val()).padStart(5, '0');
            myId = 'p1';
            isHost = true;
            localTs = Date.now();
            lastVfxLogTs = 0;
            lastVfxTurnKey = null;
            startCutscenePlayed = false;
            winCutscenePlayed = false;
            isPlayingStartCutscene = false;
            isPlayingWinCutscene = false;
            const roomRef = ref(db, `games/Escape/rooms/${currentRoomId}`);
            onDisconnect(roomRef).remove(); // Host หลุด -> ลบห้อง
            set(roomRef, {
                status: 'waiting',
                players: { p1: createPlayerObj(myName, false) },
                botCount: 0, turnIndex: 0, turnProcessed: false,
                turnFinished: null,
                log: { msg: '', ts: 0 }
            });
            playAudio('1.mp3');
            enterRoomUI();
            announceSR(`สร้างห้องสำเร็จ รหัส ${currentRoomId}`);
        }
    });
}

function joinRoom(rId) {
    currentRoomId = rId;
    isHost = false;
    lastVfxLogTs = 0;
    lastVfxTurnKey = null;
    startCutscenePlayed = false;
    winCutscenePlayed = false;
    isPlayingStartCutscene = false;
    isPlayingWinCutscene = false;
    const roomRef = ref(db, `games/Escape/rooms/${currentRoomId}`);
    let assignedSlot = null;
    runTransaction(roomRef, (rData) => {
        if (!rData || rData.status !== 'waiting') return;
        const keys = Object.keys(rData.players || {});
        if (keys.length >= 4) return;
        let slot = 'p1';
        for(let i=1; i<=4; i++) { if(!keys.includes('p'+i)) { slot = 'p'+i; break; } }
        if (!rData.players) rData.players = {};
        rData.players[slot] = createPlayerObj(myName, false);
        rData.log = { msg: `${myName} เข้าร่วมห้องสำเร็จ`, ts: Date.now(), type: 'join' };
        assignedSlot = slot;
        return rData;
    }).then((res) => {
        if (res.committed && assignedSlot) {
            myId = assignedSlot;
            localTs = 0; // รีเซ็ตเพื่อให้รับ log การเข้าร่วมห้องพร้อมทุกคน
            onDisconnect(ref(db, `games/Escape/rooms/${currentRoomId}/players/${myId}`)).update({
                isBot: true, name: "บอท_" + myName
            });
            enterRoomUI();
        } else {
            announceSR("ห้องเต็มหรือเริ่มไปแล้ว");
            currentRoomId = null;
        }
    });
}

function enterRoomUI() {
    document.getElementById('lobby-menu-section').style.display = 'none';
    document.getElementById('lobby-room-section').style.display = 'block';
    setupRoomListener();
    document.getElementById('lobby-room-title').focus();
}

window.leaveRoomAndGoHome = function() {
    stopBGM();
    clearSceneTimeouts();
    const vfxLayer = document.getElementById('et-vfx-layer');
    if (vfxLayer) vfxLayer.innerHTML = '';
    lastVfxLogTs = 0;
    lastVfxTurnKey = null;
    startCutscenePlayed = false;
    winCutscenePlayed = false;
    isPlayingStartCutscene = false;
    isPlayingWinCutscene = false;
    if (turnTimer) {
        clearTimeout(turnTimer);
        turnTimer = null;
    }
    activeTurnLock = null;
    if (currentRoomId && myId) {
        if (isHost) remove(ref(db, `games/Escape/rooms/${currentRoomId}`));
        else update(ref(db, `games/Escape/rooms/${currentRoomId}/players/${myId}`), { isBot: true, name: "บอท_"+myName });
    }
    myId = null; currentRoomId = null; isHost = false; gameState = null;
    if(roomListener) roomListener(); // unsubscribe
    document.getElementById('lobby-menu-section').style.display = 'block';
    document.getElementById('lobby-room-section').style.display = 'none';
    switchScreen('screen-welcome', 'welcome-title');
    announceSR("ออกจากห้องแล้ว");
};

function createPlayerObj(name, isBot) {
    return {
        name: name, isBot: isBot, floor: 6, hp: 200, status: 'normal',
        infection: 0, vaccine: 0, food: 0, armor: 0, gun: 0, ammo: 0,
        keys: 0, ghostKeys: 0, secretPassage: false,
        cluesFound: {},
        canUseDoor: false,
        answeringClue: null,
        actionStarted: null
    };
}

function adjustBot(delta) {
    if(!isHost || !gameState) return;
    const keys = Object.keys(gameState.players);
    const bCount = gameState.botCount || 0;
    const total = keys.length;
    
    if(delta > 0) {
        if (total >= 4) {
            announceSR("ห้องเต็ม 4 คนแล้ว ไม่สามารถเพิ่มบอทได้อีก", 'assertive');
            return;
        }
        
        // กรองชื่อบอทที่ยังว่างอยู่ (ไม่ให้ซ้ำในห้อง)
        const usedNames = keys.map(k => gameState.players[k].name);
        const availableNames = BOT_NAMES.filter(n => !usedNames.includes(n));
        
        if (availableNames.length === 0) {
            announceSR("ชื่อบอทถูกใช้งานครบหมดแล้ว ไม่สามารถเพิ่มบอทได้อีก", 'assertive');
            return;
        }
        
        // สุ่มชื่อจากรายชื่อที่ยังว่าง
        const randomName = availableNames[Math.floor(Math.random() * availableNames.length)];
        
        let slot = 'p1';
        for(let i=1; i<=4; i++) { if(!keys.includes('p'+i)) { slot = 'p'+i; break; } }
        const msg = `เพิ่มบอทสำเร็จ ได้ชื่อ ${randomName} ขณะนี้มีบอท ${bCount + 1} ตัว รวม ${total + 1} คน`;
        update(ref(db, `games/Escape/rooms/${currentRoomId}`), {
            [`players/${slot}`]: createPlayerObj(randomName, true),
            botCount: bCount + 1,
            log: { msg: msg, ts: Date.now(), type: 'bot' }
        });
    } else if(delta < 0) {
        if (bCount <= 0) {
            announceSR("ไม่มีบอทให้ลดแล้ว", 'assertive');
            return;
        }
        const bots = keys.filter(k=>gameState.players[k].isBot);
        if(bots.length > 0) {
            const targetBotKey = bots[bots.length - 1];
            const removedBotName = gameState.players[targetBotKey].name;
            const msg = `ลดบอท ${removedBotName} สำเร็จ ขณะนี้มีบอท ${bCount - 1} ตัว รวม ${total - 1} คน`;
            const updates = { 
                botCount: bCount - 1,
                log: { msg: msg, ts: Date.now(), type: 'bot' }
            };
            updates[`players/${targetBotKey}`] = null; // คืนชื่ออัตโนมัติเนื่องจากข้อมูลถูกลบ
            update(ref(db, `games/Escape/rooms/${currentRoomId}`), updates);
        }
    }
}

// Game Logic
function generateRoomsForFloor(floorNum) {
    let arr = [];
    if (floorNum >= 4) { // Floors 6, 5, 4 (20 ห้อง)
        arr = [
            {type:'key'}, {type:'ghostKey'}, {type:'secret'}, {type:'clue', cid:0}, {type:'clue', cid:1}, {type:'clue', cid:2},
            {type:'zombie'}, {type:'zombie'}, {type:'zombie'}, {type:'zombie'},
            {type:'vaccine'}, {type:'vaccine'}, {type:'vaccine'},
            {type:'food'}, {type:'food'}, {type:'food'},
            {type:'armor'}, {type:'puzzle'}, {type:'gun'}, {type:'battery'}
        ];
    } else { // Floors 3, 2, 1 (20 ห้อง ไม่มีห้องทางลับ)
        arr = [
            {type:'key'}, {type:'ghostKey'}, {type:'clue', cid:0}, {type:'clue', cid:1}, {type:'clue', cid:2},
            {type:'zombie'}, {type:'zombie'}, {type:'zombie'}, {type:'zombie'}, {type:'zombie'}, {type:'zombie'},
            {type:'vaccine'}, {type:'vaccine'}, {type:'vaccine'},
            {type:'food'}, {type:'food'},
            {type:'armor'}, {type:'puzzle'}, {type:'gun'}, {type:'battery'}
        ];
    }
    
    // Shuffle
    for(let i=arr.length-1; i>0; i--) { const j = Math.floor(Math.random()*(i+1)); [arr[i], arr[j]] = [arr[j], arr[i]]; }
    
    let clues = [];
    if(window.EscapeTowerClueBank) {
        clues = window.EscapeTowerClueBank.selectCluesForGame(floorNum);
    }
    
    let board = {};
    for(let i=0; i<20; i++) {
        board[i] = { type: arr[i].type, explored: false };
        if(arr[i].type === 'clue' && clues[arr[i].cid]) {
            board[i].clueId = clues[arr[i].cid].id;
            board[i].explored = false;
        }
    }
    return board;
}

function startGame() {
    if(!isHost || !gameState || gameState.status !== 'waiting') return;
    const pKeys = Object.keys(gameState.players || {});
    if(pKeys.length < 2) return;
    
    const btnStart = document.getElementById('btn-start-game');
    btnStart.disabled = true;
    btnStart.setAttribute('aria-disabled', 'true');
    
    stopBGM();
    
    for(let i=pKeys.length-1; i>0; i--) { const j = Math.floor(Math.random()*(i+1)); [pKeys[i], pKeys[j]] = [pKeys[j], pKeys[i]]; }
    
    const floorsData = {};
    for(let f=6; f>=1; f--) floorsData[f] = generateRoomsForFloor(f);

    update(ref(db, `games/Escape/rooms/${currentRoomId}`), {
        status: 'playing', turnIndex: 0, order: pKeys, turnProcessed: false,
        floors: floorsData, log: { msg: 'เกมเริ่มขึ้นแล้ว หาทางลงไปยังชั้น 1 และหนีออกจากตึก!', ts: Date.now() }
    });
}

let localTs = 0;
function setupRoomListener() {
    roomListener = onValue(ref(db, `games/Escape/rooms/${currentRoomId}`), (snap) => {
        gameState = snap.val();
        if(!gameState) { stopBGM(); clearSceneTimeouts(); announceSR("ห้องถูกปิด", "assertive"); setTimeout(()=>window.location.href='../index.html', 2000); return; }
        
        // Host จะจัดการเรื่องเปลี่ยนเทิร์นเมื่อได้รับสัญญาณจาก Client
        if (isHost && gameState.status === 'playing' && gameState.turnFinished && !isPlayingStartCutscene) {
            const expectedPId = gameState.order[gameState.turnIndex];
            if (gameState.turnFinished === expectedPId) {
                endTurn(expectedPId);
                return; // รอข้อมูล Snapshot ใหม่หลังจากการเปลี่ยนเทิร์นเสร็จสิ้น
            } else {
                update(ref(db, `games/Escape/rooms/${currentRoomId}`), { turnFinished: null });
            }
        }

        if(gameState.status === 'waiting') {
            startCutscenePlayed = false;
            winCutscenePlayed = false;
            isPlayingStartCutscene = false;
            isPlayingWinCutscene = false;
            lastVfxLogTs = 0;
            lastVfxTurnKey = null;
            stopBGM();

            if (gameState.log && gameState.log.ts > localTs) {
                localTs = gameState.log.ts;
                if (gameState.log.msg) {
                    if (gameState.log.type === 'join' || gameState.log.type === 'bot') {
                        playAudio('select.mp3');
                    }
                    announceSR(gameState.log.msg, 'polite');
                }
            }
            document.getElementById('lobby-room-title').textContent = `ห้อง: ${currentRoomId}`;
            const pEntries = Object.entries(gameState.players || {});
            const totalPlayers = pEntries.length;
            document.getElementById('participant-count').textContent = `${totalPlayers}/4`;
            const listEl = document.getElementById('participant-list');
            listEl.innerHTML = '';
            pEntries.forEach(([pid, p]) => {
                const avatarHtml = getParticipantAvatarHtml(pid, p);
                listEl.innerHTML += `<div class="participant-card has-js-avatar" role="listitem">${avatarHtml}<span>${escapeHtml(p.name)} ${p.isBot?'(บอท)':''}</span></div>`;
            });
            if(isHost) {
                const bCount = gameState.botCount || 0;
                
                document.getElementById('host-controls').style.display = 'block';
                document.getElementById('bot-count-display').textContent = bCount;
                
                const btnAdd = document.getElementById('btn-add-bot');
                const btnRm = document.getElementById('btn-remove-bot');
                const btnStart = document.getElementById('btn-start-game');
                
                const isMax = (totalPlayers >= 4);
                const isMin = (bCount <= 0);
                
                btnAdd.disabled = isMax;
                btnAdd.setAttribute('aria-disabled', isMax ? 'true' : 'false');
                if (isMax) {
                    btnAdd.setAttribute('aria-label', `เพิ่มบอท (คนเต็มห้อง 4 คนแล้ว ไม่สามารถเพิ่มได้)`);
                } else {
                    btnAdd.setAttribute('aria-label', `เพิ่มจำนวนบอท ปัจจุบันมี ${bCount} ตัว รวมผู้เล่น ${totalPlayers} จาก 4 คน`);
                }
                
                btnRm.disabled = isMin;
                btnRm.setAttribute('aria-disabled', isMin ? 'true' : 'false');
                if (isMin) {
                    btnRm.setAttribute('aria-label', `ลดบอท (ปัจจุบันมีบอท 0 ตัว ไม่สามารถลดได้)`);
                } else {
                    btnRm.setAttribute('aria-label', `ลดจำนวนบอท ปัจจุบันมี ${bCount} ตัว`);
                }
                
                btnStart.style.display = 'block';
                const canStart = (totalPlayers >= 2);
                btnStart.disabled = !canStart;
                btnStart.setAttribute('aria-disabled', (!canStart).toString());
            }
        } else if (gameState.status === 'playing') {
            if (!startCutscenePlayed) {
                startCutscenePlayed = true;
                isPlayingStartCutscene = true;
                if (gameState.log && gameState.log.ts > localTs) {
                    localTs = gameState.log.ts;
                    lastVfxLogTs = gameState.log.ts;
                }
                runStartCutscene(() => {
                    isPlayingStartCutscene = false;
                    if (!gameState || gameState.status !== 'playing') return;
                    if (!document.getElementById('screen-game').classList.contains('active-screen')) {
                        switchScreen('screen-game', 'game-floor-title');
                    }
                    if (gameState.log && gameState.log.msg) {
                        document.getElementById('latest-event-display').textContent = gameState.log.msg;
                    }
                    renderGameUI();
                });
                return;
            }

            if (isPlayingStartCutscene) {
                if (gameState.log && gameState.log.ts > localTs) {
                    localTs = gameState.log.ts;
                    lastVfxLogTs = gameState.log.ts;
                }
                return;
            }

            if(!document.getElementById('screen-game').classList.contains('active-screen')) {
                switchScreen('screen-game', 'game-floor-title');
            }
            
            let pendingTurnSfx = null;
            if(gameState.log && gameState.log.ts > localTs) {
                localTs = gameState.log.ts;
                if (gameState.log.sfx) {
                    pendingTurnSfx = gameState.log.sfx;
                }
                if(gameState.log.msg) {
                    announceSR(gameState.log.msg, 'polite');
                    document.getElementById('latest-event-display').textContent = gameState.log.msg;
                }
                triggerVisualEffectFromLog(gameState.log);
            }
            
            renderGameUI(pendingTurnSfx);
        } else if (gameState.status === 'ended') {
            stopBGM();
            if (turnTimer) {
                clearTimeout(turnTimer);
                turnTimer = null;
            }
            if (gameState.log && gameState.log.ts > localTs) {
                localTs = gameState.log.ts;
                if (gameState.log.sfx) {
                    playAudioSequence(gameState.log.sfx);
                }
                triggerVisualEffectFromLog(gameState.log);
            }
            if (gameState.winner && !winCutscenePlayed) {
                winCutscenePlayed = true;
                isPlayingWinCutscene = true;
                if (gameState.log && gameState.log.ts > localTs) {
                    localTs = gameState.log.ts;
                }
                const winnerName = (gameState.players && gameState.players[gameState.winner]) 
                    ? gameState.players[gameState.winner].name 
                    : 'ผู้รอดชีวิต';
                runWinCutscene(winnerName, () => {
                    isPlayingWinCutscene = false;
                    if (!gameState || gameState.status !== 'ended') return;
                    if (!document.getElementById('screen-result').classList.contains('active-screen')) {
                        switchScreen('screen-result', 'result-title');
                        renderResult();
                    }
                });
                return;
            }

            if (isPlayingWinCutscene) return;

            if(!document.getElementById('screen-result').classList.contains('active-screen')) {
                switchScreen('screen-result', 'result-title');
                renderResult();
            }
        }
    });
}

async function renderGameUI(pendingTurnSfx = null) {
    if(!gameState || isPlayingStartCutscene || isPlayingWinCutscene) return;
    const myP = gameState.players[myId];
    if(!myP) return;
    
    document.getElementById('game-floor-title').textContent = `คุณอยู่ชั้น ${myP.floor}`;
    
    const curPId = gameState.order[gameState.turnIndex];
    let curP = gameState.players[curPId];
    
    const thisTurnKey = `${gameState.turnIndex}_${curPId}`;

    // การประมวลผลช่วงเริ่มเทิร์นของ Host
    if (isHost && !gameState.turnProcessed) {
        if (curP.status === 'dead' || curP.status === 'escaped') {
            const initKey = thisTurnKey + "_init_skip";
            if (activeTurnLock === initKey) return;
            activeTurnLock = initKey;
            setTimeout(() => endTurn(curPId), 1000);
            return;
        }

        const initKey = thisTurnKey + "_init";
        if (activeTurnLock === initKey) return; 
        activeTurnLock = initKey;
        
        let turnLog = '';
        let turnSfx = [];
        
        if (curP.status === 'recovered_ready') {
            curP.hp = 150;
            curP.status = 'normal';
            turnLog += `${curP.name} พักฟื้นเสร็จสิ้น ฟื้นฟู HP เป็น 150 และพร้อมเล่นต่อ `;
            turnSfx.push('getup.mp3');
        } else if (curP.status === 'resting') {
            curP.status = 'recovered_ready';
            turnLog += `${curP.name} กำลังพักฟื้น หยุดเล่น 1 เทิร์น `;
            turnSfx.push('stop.mp3');
            await update(ref(db), { 
                [`games/Escape/rooms/${currentRoomId}/players/${curPId}`]: curP,
                [`games/Escape/rooms/${currentRoomId}/turnProcessed`]: true 
            });
            if(turnLog) await emitLog(turnLog.trim(), turnSfx.length > 0 ? turnSfx : null);
            
            setTimeout(() => {
                endTurn(curPId);
            }, 2500); 
            return;
        } else {
            if (curP.infection > 0 && curP.vaccine > 0) {
                curP.vaccine--;
                curP.infection = 0;
                turnLog += `${curP.name} ใช้วัคซีนรักษาการติดเชื้ออัตโนมัติ `;
                turnSfx.push('vac.mp3');
            }
            
            if (curP.hp > 0 && curP.hp < 200 && curP.food > 0) {
                const heal = Math.min(30, 200 - curP.hp);
                curP.food--;
                curP.hp += heal;
                turnLog += `${curP.name} กินอาหารฟื้นฟู HP ${heal} หน่วย (HP ปัจจุบัน ${curP.hp}) `;
                turnSfx.push('eat.mp3');
            }
            
            if (curP.infection > 0) {
                curP.hp -= curP.infection;
                turnLog += `${curP.name} เสีย ${curP.infection} HP จากการติดเชื้อ (เหลือ ${curP.hp} HP) `;
                turnSfx.push('reduce.mp3');
            }
            
            if (curP.hp <= 0) {
                curP.status = 'dead';
                turnLog += `${curP.name} เสียชีวิตแล้ว... `;
            } else if (curP.hp < 60 && curP.status === 'normal') {
                curP.status = 'resting';
                turnLog += `${curP.name} HP ลดต่ำกว่า 60 ได้รับบาดเจ็บสาหัส ต้องหยุดพักฟื้น `;
                turnSfx.push('stop.mp3');
            }
        }
        
        await update(ref(db), { 
            [`games/Escape/rooms/${currentRoomId}/players/${curPId}`]: curP,
            [`games/Escape/rooms/${currentRoomId}/turnProcessed`]: true 
        });
        
        if (turnLog) {
            await emitLog(turnLog.trim(), turnSfx.length > 0 ? turnSfx : null);
            await new Promise(resolve => setTimeout(resolve, 2700));
        }
        
        if (curP.status !== 'normal') {
            setTimeout(() => {
                endTurn(curPId);
            }, 3200); 
            return;
        }
        // รอรับ Snapshot ล่าสุดก่อนปลดปล่อยให้เกมดำเนินต่อไป
        return; 
    }
    
    if (!gameState.turnProcessed && isHost) return;

    curP = gameState.players[curPId]; 
    const isMyTurn = (curPId === myId && !myP.isBot && myP.status === 'normal' && gameState.turnProcessed);
    
    // ตั้งค่าตัวจับเวลาเมื่อมีการเปลี่ยนเทิร์นสมบูรณ์
    if(lastAnnouncedTurn !== thisTurnKey) {
        lastAnnouncedTurn = thisTurnKey;
        activeTurnLock = null; // ปลดล็อก UI สำหรับผู้เล่นเมื่อเริ่มเทิร์นใหม่
        
        announceSR(`ถึงเทิร์นของ ${curP.name}`);
        triggerTurnChangeVisual(thisTurnKey, curPId, curP.name);

        let startTurnAudio = ['turn.mp3'];
        if (pendingTurnSfx) {
            const sfxArr = Array.isArray(pendingTurnSfx) ? pendingTurnSfx : [pendingTurnSfx];
            startTurnAudio.push(...sfxArr);
            pendingTurnSfx = null;
        }
        if (curPId === myId && !myP.isBot && myP.status === 'normal') {
            startTurnAudio.push('abc.mp3');
        }
        playAudioSequence(startTurnAudio);
        
        if (turnTimer) {
            clearTimeout(turnTimer);
            turnTimer = null;
        }

        if(isHost && curP.isBot && curP.status === 'normal') {
            setTimeout(() => executeBotTurn(curPId, thisTurnKey), 3500); 
        } else if (curP.status !== 'normal' && isHost) {
             const skipKey = thisTurnKey + "_skip";
             if (activeTurnLock !== skipKey) {
                 activeTurnLock = skipKey;
                 setTimeout(() => endTurn(curPId), 2000);
             }
        } else if (isHost && !curP.isBot && curP.status === 'normal') {
            currentTurnKey = thisTurnKey;
            turnTimer = setTimeout(async () => {
                if (!gameState || gameState.status !== 'playing') return;
                if (currentTurnKey !== thisTurnKey) return; 
                if (activeTurnLock === thisTurnKey) return; // หากผู้เล่นกำลังทำกิจกรรมอยู่ให้ยกเลิกการหมดเวลา
                
                let p = gameState.players[curPId];
                if (p.actionStarted && (Date.now() - p.actionStarted < 10000)) return; 
                
                activeTurnLock = thisTurnKey; // ยึดล็อกสำหรับกระบวนการหมดเวลา
                
                if (p.answeringClue) {
                    await emitLog(`${p.name} หมดเวลาตอบคำใบ้ข้ามเทิร์นทันที`);
                    await update(ref(db), { [`games/Escape/rooms/${currentRoomId}/players/${curPId}/answeringClue`]: null });
                    setTimeout(() => {
                        endTurn(curPId);
                    }, 2000);
                } else {
                    await emitLog(`${p.name} หมดเวลาดำเนินการ ระบบเล่นอัตโนมัติสำหรับเทิร์นนี้`);
                    setTimeout(() => {
                        activeTurnLock = null; // คืนล็อกเพื่อให้บอทนำไปประมวลผลต่อ
                        executeBotTurn(curPId, thisTurnKey);
                    }, 2000);
                }
            }, TURN_TIME_LIMIT);
        }
    }

    if (pendingTurnSfx) {
        playAudioSequence(pendingTurnSfx);
    }
    
    // บังคับปิด Modal ทันที หากหมดเวลาและเทิร์นเปลี่ยนไปแล้ว
    const modal = document.getElementById('clue-modal');
    if (modal.style.display !== 'none' && !isMyTurn) {
        modal.style.display = 'none';
        document.getElementById('clue-answer-input').disabled = true;
        document.getElementById('btn-submit-clue').disabled = true;
    }
    
    document.getElementById('game-status-bar').textContent = `เทิร์นของ: ${curP.name}`;
    
    const statusText = `สถานะของคุณ: ${myP.name} ชั้น ${myP.floor} HP ${myP.hp} อาหารในคลัง ${myP.food} วัคซีน ${myP.vaccine} กุญแจ ${myP.keys} ปืน ${myP.gun} กระสุน ${myP.ammo}`;
    const statusContainer = document.getElementById('my-player-status');
    statusContainer.setAttribute('aria-label', statusText);
    if (curPId === myId) {
        statusContainer.classList.add('et-my-turn-glow');
    } else {
        statusContainer.classList.remove('et-my-turn-glow');
    }
    const myAvatarHtml = getParticipantAvatarHtml(myId, myP);
    const safeMyName = escapeHtml(myP.name);
    statusContainer.innerHTML = `${myAvatarHtml}<span class="et-status-bar-items" aria-hidden="true">` +
        `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">👤</span>${safeMyName}</span>` +
        `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">🏢</span>ชั้น&nbsp;${myP.floor}</span>` +
        `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">❤️</span>HP:&nbsp;${myP.hp}</span>` +
        `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">🥫</span>อาหาร:&nbsp;${myP.food}</span>` +
        `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">💉</span>วัคซีน:&nbsp;${myP.vaccine}</span>` +
        `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">🔑</span>กุญแจ:&nbsp;${myP.keys}</span>` +
        `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">🔫</span>ปืน:&nbsp;${myP.gun}&nbsp;<span class="et-stat-icon" aria-hidden="true">🔋</span>(${myP.ammo})</span>` +
        `</span>`;
    
    const cardsEl = document.getElementById('player-status-cards');
    cardsEl.innerHTML = '';
    gameState.order.forEach(pid => {
        const p = gameState.players[pid];
        let cClass = 'p-status-card';
        if(pid === curPId) cClass += ' active-turn';
        if(p.status === 'dead') cClass += ' dead';
        if(p.status === 'escaped') cClass += ' escaped';
        const pAvatarHtml = getParticipantAvatarHtml(pid, p);
        const safePName = escapeHtml(p.name);
        const pRoleIcon = p.isBot ? '🤖' : '👤';
        cardsEl.innerHTML += `<div class="${cClass}" aria-hidden="true">
            ${pAvatarHtml}
            <strong><span class="et-stat-icon" aria-hidden="true">${pRoleIcon}</span> ${safePName} (<span class="et-stat-icon" aria-hidden="true">🏢</span> ชั้น ${p.floor})</strong><br>
            <span class="et-stat-icon" aria-hidden="true">❤️</span> HP: ${p.hp} | <span class="et-stat-icon" aria-hidden="true">☣️</span> ติดเชื้อ: ${p.infection}<br>
            <span class="et-stat-icon" aria-hidden="true">🔑</span> กุญแจ: ${p.keys} | <span class="et-stat-icon" aria-hidden="true">🥫</span> อาหาร: ${p.food} | <span class="et-stat-icon" aria-hidden="true">🛡️</span> เกราะ: ${p.armor} | <span class="et-stat-icon" aria-hidden="true">🔫</span> ปืน: ${p.gun} <span class="et-stat-icon" aria-hidden="true">🔋</span>(${p.ammo})
        </div>`;
    });
    
    const boardEl = document.getElementById('escape-board');
    boardEl.innerHTML = '';
    const floorData = gameState.floors[myP.floor];
    if(floorData) {
        for(let i=0; i<20; i++) {
            const roomData = floorData[i];
            const btn = document.createElement('button');
            btn.className = 'board-cell';
            const num = i + 1;
            const isClue = isClueRoomType(roomData.type);

            // ข้อยกเว้น: ห้องคำใบ้ต้องแสดงเป็นประตูปิดเสมอ และสำรวจซ้ำได้ตามกติกาเดิม
            // ในโค้ดจริง ห้องคำใบ้ไม่ได้ถูกตั้งค่า explored = true ใน Firebase แต่รองรับไว้ทั้งกรณี未สำรวจและกรณี explored = true
            if (isClue) {
                btn.classList.add('room-clue-door');
                btn.textContent = num.toString();
                if (roomData.explored) {
                    btn.classList.add('explored');
                    btn.setAttribute('aria-label', `ห้องคำใบ้ ${num} สำรวจแล้ว สามารถกดอ่านคำใบ้ซ้ำได้`);
                } else {
                    btn.setAttribute('aria-label', `ห้อง ${num} กดเพื่อสำรวจ`);
                }
                if (isMyTurn && activeTurnLock !== thisTurnKey) {
                    btn.onclick = () => exploreRoom(i, roomData);
                } else {
                    btn.disabled = true;
                    btn.setAttribute('aria-disabled', 'true');
                }
            } else if (roomData.explored) {
                btn.classList.add('explored');
                const vInfo = ROOM_VISUAL_MAP[roomData.type];
                if (vInfo) {
                    btn.classList.add(vInfo.cls);
                    btn.innerHTML = `<span class="room-num-label" aria-hidden="true">${num}</span><span class="room-result-icon" aria-hidden="true">${vInfo.icon}</span>`;
                } else {
                    btn.textContent = num.toString();
                }
                btn.disabled = true;
                btn.setAttribute('aria-disabled', 'true');
                btn.setAttribute('aria-label', `ห้อง ${num} สำรวจแล้ว`);
            } else {
                btn.textContent = num.toString();
                btn.setAttribute('aria-label', `ห้อง ${num} กดเพื่อสำรวจ`);
                if(isMyTurn && activeTurnLock !== thisTurnKey) {
                    btn.onclick = () => exploreRoom(i, roomData);
                } else {
                    btn.disabled = true;
                    btn.setAttribute('aria-disabled', 'true');
                }
            }
            boardEl.appendChild(btn);
        }
    }
    
    const btnStairs = document.getElementById('btn-use-stairs');
    const isFloor1 = (myP.floor === 1);
    
    if (isFloor1) {
        btnStairs.textContent = "🚪 เปิดประตูหนีออกจากตึก";
    } else {
        btnStairs.textContent = "🚪 ลงบันไดไปชั้นล่าง";
    }
    
    const canUseStairs = isMyTurn && activeTurnLock !== thisTurnKey && (myP.keys > 0 || myP.ghostKeys > 0 || myP.secretPassage || myP.canUseDoor);
    btnStairs.disabled = !canUseStairs;
    btnStairs.setAttribute('aria-disabled', (!canUseStairs).toString());
    
    if (canUseStairs) {
        if (isFloor1) {
            btnStairs.setAttribute('aria-label', "เปิดประตูหนีออกจากตึก คุณมีไอเท็มพร้อมแล้ว กดเพื่อหนีรอดและชนะเกม");
        } else {
            btnStairs.setAttribute('aria-label', "ลงบันไดไปชั้นล่าง คุณมีไอเท็มพร้อมลงชั้นแล้ว กดเพื่อลง");
        }
        btnStairs.onclick = useStairs;
    } else {
        if (isFloor1) {
            btnStairs.setAttribute('aria-label', "ประตูหนีออกจากตึก (ยังเปิดไม่ได้ ต้องหากุญแจ ทางลับ หรือตอบคำใบ้ก่อน)");
        } else {
            btnStairs.setAttribute('aria-label', "บันไดลงชั้นล่าง (ยังลงไม่ได้ ต้องหากุญแจ ทางลับ หรือตอบคำใบ้ก่อน)");
        }
    }
}

function readAllStatus() {
    if(!gameState) return;
    const modal = document.getElementById('all-status-modal');
    const list = document.getElementById('all-status-list');
    list.innerHTML = '';
    const curPId = (gameState.order && gameState.order[gameState.turnIndex]) ? gameState.order[gameState.turnIndex] : null;
    gameState.order.forEach(pid => {
        const p = gameState.players[pid];
        let txt = `${p.name} - ชั้น ${p.floor} | HP: ${p.hp}`;
        if(p.status !== 'normal') txt += ` (${p.status})`;
        txt += ` | ติดเชื้อ: ${p.infection} | อาหาร: ${p.food} | วัคซีน: ${p.vaccine} | กุญแจ: ${p.keys} | เกราะ: ${p.armor} | ปืน: ${p.gun} | กระสุน: ${p.ammo}`;
        const avatarHtml = getParticipantAvatarHtml(pid, p);
        const safeName = escapeHtml(p.name);
        const safeSrTxt = escapeHtml(txt);
        const roleIcon = p.isBot ? '🤖' : '👤';
        let statusBadgeHtml = '';
        if (p.status !== 'normal') {
            let stateIcon = '⚠️';
            if (p.status === 'dead') stateIcon = '💀';
            else if (p.status === 'escaped') stateIcon = '🏆';
            else if (p.status === 'resting' || p.status === 'recovered_ready') stateIcon = '🛌';
            statusBadgeHtml = `<span class="et-state-tag"><span class="et-stat-icon" aria-hidden="true">${stateIcon}</span>(${escapeHtml(p.status)})</span>`;
        }
        const activeTurnBadgeHtml = (pid === curPId)
            ? `<span class="et-stat-pill" style="border-color:#fbbf24!important;color:#fde047!important;" aria-hidden="true">⚡ เทิร์นปัจจุบัน</span>`
            : '';
        const visualHtml = `<div class="et-all-status-body" aria-hidden="true">` +
            `<div class="et-all-status-header">` +
                `<span><span class="et-stat-icon" aria-hidden="true">${roleIcon}</span> ${safeName}</span>` +
                `${activeTurnBadgeHtml}` +
                `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">🏢</span>ชั้น ${p.floor}</span>` +
                `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">❤️</span>HP: ${p.hp}</span>` +
                `${statusBadgeHtml}` +
            `</div>` +
            `<div class="et-all-status-grid">` +
                `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">☣️</span>ติดเชื้อ: ${p.infection}</span>` +
                `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">🥫</span>อาหาร: ${p.food}</span>` +
                `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">💉</span>วัคซีน: ${p.vaccine}</span>` +
                `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">🔑</span>กุญแจ: ${p.keys}</span>` +
                `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">🛡️</span>เกราะ: ${p.armor}</span>` +
                `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">🔫</span>ปืน: ${p.gun}</span>` +
                `<span class="et-stat-pill"><span class="et-stat-icon" aria-hidden="true">🔋</span>กระสุน: ${p.ammo}</span>` +
            `</div>` +
        `</div>`;
        list.innerHTML += `<div class="participant-card has-js-avatar" role="listitem">${avatarHtml}<span class="sr-only">${safeSrTxt}</span>${visualHtml}</div>`;
    });
    modal.style.display = 'flex';
    document.getElementById('screen-game').setAttribute('aria-hidden', 'true');
    const title = document.getElementById('all-status-title');
    title.setAttribute('tabindex', '-1');
    title.focus();
    announceSR("เปิดหน้าสถานะตัวละครทั้งหมดแล้ว");
}

function closeAllStatus() {
    document.getElementById('all-status-modal').style.display = 'none';
    document.getElementById('screen-game').removeAttribute('aria-hidden');
    document.getElementById('btn-read-all-status').focus();
    announceSR("ปิดหน้าสถานะแล้ว กลับสู่หน้าเกม");
}

async function exploreRoom(roomIdx, roomData) {
    if(!gameState) return;
    const thisTurnKey = `${gameState.turnIndex}_${myId}`;
    if (activeTurnLock === thisTurnKey) return;
    activeTurnLock = thisTurnKey;
    
    const pId = myId;
    let p = gameState.players[pId];

    if(roomData.type === 'clue') {
        update(ref(db), { [`games/Escape/rooms/${currentRoomId}/players/${pId}/answeringClue`]: true }).then(async () => {
            const isOpen = openClueModal(roomData.clueId);
            if (isOpen) {
                activeTurnLock = null; // คืนล็อกชั่วคราวให้ผู้เล่นตอบคำถามได้
                await emitLog('', ['question.mp3']);
            }
        });
        return;
    }
    
    await update(ref(db), { 
        [`games/Escape/rooms/${currentRoomId}/floors/${p.floor}/${roomIdx}/explored`]: true,
        [`games/Escape/rooms/${currentRoomId}/players/${pId}/actionStarted`]: Date.now()
    });
    
    let msg = `${p.name} สำรวจห้อง ${roomIdx+1}: `;
    let sfxList = [];
    
    switch(roomData.type) {
        case 'key': p.keys++; msg += "พบกุญแจปกติ 1 ดอก"; sfxList.push('key1.mp3'); break;
        case 'ghostKey': p.ghostKeys++; msg += "พบกุญแจผี 1 ดอก"; sfxList.push('key2.mp3'); break;
        case 'secret': p.secretPassage = true; msg += "พบทางลงลับ! สามารถลงข้ามชั้นได้"; sfxList.push('key3.mp3'); break;
        case 'vaccine': 
            p.vaccine++; 
            msg += "พบวัคซีน 1 ขวด เก็บเข้าคลัง"; 
            sfxList.push('findvac.mp3');
            break;
        case 'food': 
            sfxList.push('findfood.mp3');
            if(p.hp < 200) {
                const heal = Math.min(30, 200 - p.hp);
                p.hp += heal;
                msg += `พบอาหาร กินทันทีฟื้นฟู HP ${heal} หน่วย (HP ปัจจุบัน ${p.hp})`;
                sfxList.push('eat.mp3');
            } else {
                p.food = (p.food || 0) + 1;
                msg += `พบอาหาร แต่ HP เต็ม 200 แล้ว จึงเก็บอาหารเข้าคลัง (มีอาหารสะสม ${p.food} ชิ้น)`;
            }
            break;
        case 'armor': 
            p.armor++; 
            msg += "พบเกราะสวรรค์ 1 ชิ้น"; 
            sfxList.push('findshield.mp3');
            break;
        case 'battery': 
            sfxList.push('findbat.mp3');
            if(p.ammo < 2) { p.ammo++; msg += "พบแบตเตอรี่ ได้กระสุน 1 นัด"; }
            else { msg += "พบแบตเตอรี่ แต่กระสุนเต็มแล้ว"; }
            break;
        case 'gun':
            sfxList.push('findgun.mp3');
            if(p.gun === 0) { p.gun = 1; if(p.ammo < 2) p.ammo++; msg += "พบปืนเลเซอร์และกระสุน 1 นัด!"; }
            else if(p.ammo < 2) { p.ammo++; msg += "พบปืนเลเซอร์ ซ้ำ แปลงเป็นกระสุน 1 นัด"; }
            else { msg += "พบปืนเลเซอร์ ซ้ำ แต่กระสุนเต็มแล้ว"; }
            break;
        case 'puzzle': 
            msg += "พบห้องน้ำเก่าสกปรกแถมผุพัง: อั้นไว้ก่อนนะ"; 
            sfxList.push('toilet.mp3');
            break;
        case 'zombie':
            msg += "พบซอมบี้! ";
            if(p.gun > 0 && p.ammo > 0) {
                p.ammo--; msg += "ใช้ปืนยิงป้องกันสำเร็จ เสียกระสุน 1 นัด";
                sfxList.push('zombi2.mp3');
            } else if (p.armor > 0) {
                p.armor--; msg += "เกราะสวรรค์ป้องกันการโจมตีสำเร็จ เกราะพัง";
                sfxList.push('zombi3.mp3');
            } else {
                if(p.infection === 0) p.infection = 20;
                else p.infection = 40;
                msg += `ถูกโจมตี ติดเชื้อระดับ ${p.infection} HP/เทิร์น`;
                sfxList.push('zombi1.mp3');
            }
            break;
    }
    
    if(p.hp <= 0) {
        p.status = 'dead';
        msg += ` | ${p.name} เสียชีวิตแล้ว...`;
    } else if(p.hp < 60 && p.status === 'normal') {
        p.status = 'resting';
        msg += ` | HP ลดต่ำกว่า 60 (${p.hp} HP) ต้องหยุดพักฟื้นในเทิร์นถัดไป`;
    }
    
    await emitLog(msg, sfxList.length > 0 ? sfxList : null);
    await commitPlayer(pId, p);
    
    setTimeout(() => {
        endTurn(pId);
    }, 3200); 
}

function openClueModal(cid) {
    const clue = window.EscapeTowerClueBank ? window.EscapeTowerClueBank.getClueById(cid) : null;
    if(!clue) {
        emitLog(`${gameState.players[myId].name} สำรวจห้องคำใบ้ แต่ระบบคำใบ้ไม่พร้อมใช้งาน รับสิทธิ์ผ่านประตูไปเลย`, ['question.mp3']);
        gameState.players[myId].canUseDoor = true;
        
        update(ref(db), { 
            [`games/Escape/rooms/${currentRoomId}/players/${myId}/answeringClue`]: null,
            [`games/Escape/rooms/${currentRoomId}/players/${myId}/actionStarted`]: Date.now()
        });
        
        commitPlayer(myId, gameState.players[myId]).then(() => setTimeout(()=> { 
            endTurn(myId); 
        }, 3200));
        return false;
    }
    
    const m = document.getElementById('clue-modal');
    document.getElementById('clue-question').textContent = clue.question;
    
    const inputEl = document.getElementById('clue-answer-input');
    inputEl.value = '';
    inputEl.dataset.cid = cid;
    inputEl.disabled = false;
    document.getElementById('btn-submit-clue').disabled = false;
    
    document.getElementById('clue-feedback').textContent = '';
    m.style.display = 'flex';
    document.getElementById('clue-title').focus();
    announceSR(clue.question);
    return true;
}

async function submitClue() {
    const m = document.getElementById('clue-modal');
    if (m.style.display === 'none') return;
    
    const inputEl = document.getElementById('clue-answer-input');
    const btnSubmit = document.getElementById('btn-submit-clue');
    
    if (inputEl.disabled) return;
    
    const cid = inputEl.dataset.cid;
    const ans = inputEl.value.trim();
    if(!ans) return;
    
    if (!gameState) return;
    const thisTurnKey = `${gameState.turnIndex}_${myId}`;
    if (activeTurnLock === thisTurnKey) return;
    activeTurnLock = thisTurnKey;
    
    inputEl.disabled = true;
    btnSubmit.disabled = true;
    
    await update(ref(db), { 
        [`games/Escape/rooms/${currentRoomId}/players/${myId}/answeringClue`]: null,
        [`games/Escape/rooms/${currentRoomId}/players/${myId}/actionStarted`]: Date.now()
    });
    
    const isCorrect = window.EscapeTowerClueBank.checkAnswer(cid, ans);
    if(isCorrect) {
        m.style.display = 'none';
        const p = gameState.players[myId];
        p.canUseDoor = true;
        const targetDesc = (p.floor === 1) ? "เปิดประตูทางออกตึก" : `เปิดประตูชั้น ${p.floor}`;
        emitLog(`${p.name} ตอบคำใบ้ถูกต้อง ได้รับสิทธิ์${targetDesc}`, ['TRUE.mp3']);
        commitPlayer(myId, p).then(() => setTimeout(()=> { endTurn(myId); }, 2500));
    } else {
        document.getElementById('clue-feedback').textContent = "คำตอบผิด ลองใหม่ในเทิร์นหน้า";
        announceSR("คำตอบผิด ลองใหม่ในเทิร์นหน้า");
        emitLog('', ['wrong.mp3']);
        setTimeout(() => {
            m.style.display = 'none';
            emitLog(`${gameState.players[myId].name} ตอบคำใบ้ผิด ประตูยังไม่เปิด`);
            setTimeout(() => { endTurn(myId); }, 2000);
        }, 2000);
    }
}

async function useStairs() {
    if(!gameState) return;
    const thisTurnKey = `${gameState.turnIndex}_${myId}`;
    if (activeTurnLock === thisTurnKey) return;
    activeTurnLock = thisTurnKey;
    
    await update(ref(db), { [`games/Escape/rooms/${currentRoomId}/players/${myId}/actionStarted`]: Date.now() });
    let p = gameState.players[myId];

    let used = "";
    if(p.secretPassage) { p.secretPassage = false; used = "ทางลับ"; }
    else if(p.canUseDoor) { p.canUseDoor = false; used = "รหัสคำใบ้"; }
    else if(p.ghostKeys > 0) { p.ghostKeys--; used = "กุญแจผี"; }
    else if(p.keys > 0) { p.keys--; used = "กุญแจปกติ"; }

    if(p.floor === 1) {
        p.status = 'escaped';
        p.canUseDoor = false;
        let msg = `${p.name} ใช้${used} ปลดล็อกประตูทางออกชั้น 1 และหนีรอดออกจากตึกได้สำเร็จ เป็นผู้ชนะ!!`;
        await emitLog(msg);
        await commitPlayer(myId, p);
        await triggerEndGame(myId);
        return;
    }

    let down = (used === "ทางลับ") ? 2 : 1;
    p.floor = Math.max(1, p.floor - down);
    p.canUseDoor = false;
    
    let msg = `${p.name} ใช้${used} ลงมายังชั้น ${p.floor}`;
    if(p.floor === 1) {
        msg += " ถึงชั้น 1 แล้ว! จงหากุญแจหรือรหัสเพื่อปลดล็อกประตูหนีออกจากตึก!";
    }
    
    await emitLog(msg, ['door1.mp3']);
    await commitPlayer(myId, p);
    
    setTimeout(() => { endTurn(myId); }, 2500);
}

async function executeBotTurn(botId, turnKey) {
    if(!gameState) return;
    const tk = turnKey || `${gameState.turnIndex}_${botId}`;
    if (activeTurnLock === tk) return;
    activeTurnLock = tk;
    
    let p = gameState.players[botId];
    
    const canUnlock = (p.secretPassage || p.ghostKeys > 0 || p.keys > 0 || p.canUseDoor);
    
    if(p.floor === 1 && canUnlock) {
        let used = p.secretPassage ? "ทางลับ" : (p.ghostKeys > 0 ? "กุญแจผี" : "กุญแจปกติ");
        if(p.secretPassage) p.secretPassage = false;
        else if(p.ghostKeys > 0) p.ghostKeys--;
        else if(p.keys > 0) p.keys--;
        p.canUseDoor = false;
        
        p.status = 'escaped';
        let msg = `${p.name} ใช้${used} ปลดล็อกประตูทางออกชั้น 1 หนีรอดออกจากตึกสำเร็จ เป็นผู้ชนะ!!`;
        await emitLog(msg);
        await commitPlayer(botId, p);
        await triggerEndGame(botId);
        return;
    }
    
    if(p.floor > 1 && canUnlock) {
        let down = p.secretPassage ? 2 : 1;
        let used = p.secretPassage ? "ทางลับ" : (p.ghostKeys > 0 ? "กุญแจผี" : "กุญแจปกติ");
        if(p.secretPassage) p.secretPassage = false;
        else if(p.ghostKeys > 0) p.ghostKeys--;
        else if(p.keys > 0) p.keys--;
        p.canUseDoor = false;
        
        p.floor = Math.max(1, p.floor - down);
        let msg = `${p.name} ใช้${used} ลงมายังชั้น ${p.floor}`;
        if(p.floor === 1) {
            msg += " ถึงชั้น 1 แล้ว! กำลังหาทางเปิดประตูหนีออกจากตึก";
        }
        await emitLog(msg, ['door1.mp3']);
        await commitPlayer(botId, p);
        setTimeout(() => { endTurn(botId); }, 2500);
        return;
    }
    
    const floorData = gameState.floors[p.floor];
    let available = [];
    for(let i=0; i<20; i++) {
        if(!floorData[i].explored && floorData[i].type !== 'clue') available.push(i);
    }
    
    if(available.length > 0) {
        const choice = available[Math.floor(Math.random()*available.length)];
        const roomData = floorData[choice];
        await update(ref(db), { [`games/Escape/rooms/${currentRoomId}/floors/${p.floor}/${choice}/explored`]: true });
        
        let msg = `${p.name} สำรวจห้อง ${choice+1}: `;
        let sfxList = [];
        switch(roomData.type) {
            case 'key': p.keys++; msg += "พบกุญแจปกติ"; sfxList.push('key1.mp3'); break;
            case 'ghostKey': p.ghostKeys++; msg += "พบกุญแจผี"; sfxList.push('key2.mp3'); break;
            case 'secret': p.secretPassage = true; msg += "พบทางลับ"; sfxList.push('key3.mp3'); break;
            case 'vaccine': 
                p.vaccine++; 
                msg += "พบวัคซีน"; 
                sfxList.push('findvac.mp3');
                break;
            case 'food': 
                sfxList.push('findfood.mp3');
                if(p.hp < 200) {
                    p.hp = Math.min(200, p.hp + 30);
                    msg += "กินอาหารฟื้น HP";
                    sfxList.push('eat.mp3');
                } else {
                    p.food = (p.food || 0) + 1;
                    msg += "เก็บอาหารเข้าคลังสำรอง";
                }
                break;
            case 'armor': 
                p.armor++; 
                msg += "ได้เกราะ"; 
                sfxList.push('findshield.mp3');
                break;
            case 'battery': 
                sfxList.push('findbat.mp3');
                if(p.ammo<2) p.ammo++; 
                msg += "ได้กระสุน"; 
                break;
            case 'gun': 
                sfxList.push('findgun.mp3');
                if(p.gun===0) p.gun=1; 
                if(p.ammo<2) p.ammo++; 
                msg += "ได้ปืนและกระสุน"; 
                break;
            case 'puzzle': 
                msg += "พบห้องน้ำเก่าสกปรกแถมผุพัง: อั้นไว้ก่อนนะ"; 
                sfxList.push('toilet.mp3');
                break;
            case 'zombie':
                msg += "เจอซอมบี้! ";
                if(p.gun>0 && p.ammo>0) { 
                    p.ammo--; 
                    msg+="ยิงป้องกันสำเร็จ"; 
                    sfxList.push('zombi2.mp3');
                }
                else if(p.armor>0) { 
                    p.armor--; 
                    msg+="เกราะกันสำเร็จ"; 
                    sfxList.push('zombi3.mp3');
                }
                else { 
                    p.infection = p.infection===0?20:40; 
                    msg+="ติดเชื้อ!"; 
                    sfxList.push('zombi1.mp3');
                }
                break;
        }
        
        if(p.hp <= 0) {
            p.status = 'dead';
            msg += ` | ${p.name} เสียชีวิตแล้ว...`;
        } else if(p.hp > 0 && p.hp < 60 && p.status === 'normal') {
            p.status = 'resting';
            msg += ` | HP ลดต่ำกว่า 60 (${p.hp} HP) ต้องหยุดพักฟื้นในเทิร์นถัดไป`;
        }

        await emitLog(msg, sfxList.length > 0 ? sfxList : null);
        await commitPlayer(botId, p);
    } else {
        await emitLog(`${p.name} ไม่เหลือห้องให้สำรวจ ข้ามเทิร์น`);
    }
    setTimeout(() => { endTurn(botId); }, 2500);
}

async function triggerEndGame(winnerId) {
    const roomRef = ref(db, `games/Escape/rooms/${currentRoomId}`);
    await runTransaction(roomRef, (rData) => {
        if (rData && rData.status === 'playing') {
            rData.status = 'ended';
            rData.winner = winnerId;
            return rData;
        }
        return;
    });
}

async function emitLog(msg, sfx = null) {
    const logData = { msg, ts: Date.now() };
    if (sfx) logData.sfx = sfx;
    await update(ref(db), { [`games/Escape/rooms/${currentRoomId}/log`]: logData });
}

async function commitPlayer(pId, pData) {
    await update(ref(db), { [`games/Escape/rooms/${currentRoomId}/players/${pId}`]: pData });
}

function endTurn(curPId) {
    if (!gameState || gameState.status === 'ended') return;
    
    let allDead = Object.values(gameState.players).every(p => p.status === 'dead');
    if (allDead && isHost) {
        update(ref(db, `games/Escape/rooms/${currentRoomId}`), { status: 'ended' });
        return;
    }

    if (isHost) {
        const expectedPId = gameState.order[gameState.turnIndex];
        if (expectedPId !== curPId) return; // ป้องกัน Double turn advance หากมี callback ทำงานซ้อน
        
        if (turnTimer) {
            clearTimeout(turnTimer);
            turnTimer = null;
        }

        const order = gameState.order;
        let idx = gameState.turnIndex;
        idx = (idx + 1) % order.length;
        
        update(ref(db), { 
            [`games/Escape/rooms/${currentRoomId}/turnIndex`]: idx,
            [`games/Escape/rooms/${currentRoomId}/turnProcessed`]: false,
            [`games/Escape/rooms/${currentRoomId}/turnFinished`]: null,
            [`games/Escape/rooms/${currentRoomId}/players/${curPId}/actionStarted`]: null,
            [`games/Escape/rooms/${currentRoomId}/players/${curPId}/answeringClue`]: null
        });
    } else {
        const expectedPId = gameState.order[gameState.turnIndex];
        if (expectedPId === curPId && gameState.turnFinished !== curPId) {
            update(ref(db, `games/Escape/rooms/${currentRoomId}`), { turnFinished: curPId });
        }
    }
}

function renderResult() {
    const winnerName = gameState.winner ? gameState.players[gameState.winner].name : "ไม่มีผู้รอดชีวิต";
    
    document.getElementById('result-title').textContent = `ผู้ชนะ: ${winnerName}`;
    
    const resEl = document.getElementById('result-list');
    resEl.innerHTML = '';
    const pArr = Object.entries(gameState.players).map(([id, p]) => ({ id, ...p }));
    
    // เรียงลำดับจาก: 1. ผู้ชนะ (escaped) -> 2. ผู้ที่ยังรอดชีวิตเรียงตามชั้นที่ใกล้ชั้น 1 ที่สุด -> 3. ผู้ที่เสียชีวิตเรียงตามชั้นที่ใกล้ชั้น 1 ที่สุด
    pArr.sort((a, b) => {
        const aIsWinner = (a.status === 'escaped' || a.id === gameState.winner);
        const bIsWinner = (b.status === 'escaped' || b.id === gameState.winner);
        if (aIsWinner && !bIsWinner) return -1;
        if (!aIsWinner && bIsWinner) return 1;
        
        const aIsDead = (a.status === 'dead');
        const bIsDead = (b.status === 'dead');
        if (!aIsDead && bIsDead) return -1;
        if (aIsDead && !bIsDead) return 1;
        
        // เรียงตามชั้นที่ใกล้ชั้น 1 มากที่สุด (เลขชั้นน้อยกว่าอยู่อันดับดีกว่า)
        if (a.floor !== b.floor) return a.floor - b.floor;
        
        // หากอยู่ชั้นเดียวกัน เรียงตามผู้ที่มีสิทธิ์เปิดประตู/กุญแจพร้อมกว่า
        const aReady = (a.keys > 0 || a.ghostKeys > 0 || a.secretPassage || a.canUseDoor) ? 1 : 0;
        const bReady = (b.keys > 0 || b.ghostKeys > 0 || a.secretPassage || b.canUseDoor) ? 1 : 0;
        if (aReady !== bReady) return bReady - aReady;
        
        // หากเท่ากันอีก เรียงตาม HP มากที่สุด
        return b.hp - a.hp;
    });

    let summarySpeech = `เกมจบลงแล้ว ผู้ชนะคือ ${winnerName} สรุปอันดับดังนี้ `;
    pArr.forEach((p, idx) => {
        const rank = idx + 1;
        let txt = `อันดับที่ ${rank}: ${p.name} - ชั้น ${p.floor} `;
        if (p.status === 'escaped' || p.id === gameState.winner) {
            txt += `(หนีรอดสำเร็จ! 🏆 HP: ${p.hp})`;
        } else if (p.status === 'dead') {
            txt += `(เสียชีวิต 💀 HP: 0)`;
        } else {
            txt += `(ติดอยู่ในตึก HP: ${p.hp})`;
        }
        const avatarHtml = getParticipantAvatarHtml(p.id, p);
        resEl.innerHTML += `<div class="participant-card has-js-avatar" role="listitem">${avatarHtml}<span>${escapeHtml(txt)}</span></div>`;
        summarySpeech += `${txt} `;
    });
    
    announceSR(summarySpeech.trim(), "assertive");
}