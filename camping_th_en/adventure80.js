import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
import { getDatabase, ref, set, onValue, update, runTransaction, remove, onDisconnect } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-database.js";

// Exact Firebase Configuration
const firebaseConfig = {
    apiKey: "AIzaSyDvcdgsyT5sDdYTYKIqetzNL9Be-MFC0l4",
    authDomain: "xo-game-134ec.firebaseapp.com",
    databaseURL: "https://xo-game-134ec-default-rtdb.asia-southeast1.firebasedatabase.app",
    projectId: "xo-game-134ec",
    storageBucket: "xo-game-134ec.firebasestorage.app",
    messagingSenderId: "318375224157",
    appId: "1:318375224157:web:8dfa7c16557b2890b77eb4"
};

const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

// ==================================================
// BILINGUAL SUPPORT LAYER (TH Default 100% Preserved / EN)
// ==================================================
function detectBrowserLanguage() {
    try {
        const saved = localStorage.getItem('adventure80_lang');
        if (saved === 'th' || saved === 'en') return saved;
    } catch (e) {}
    const lang = (navigator.language || navigator.userLanguage || 'en').toLowerCase();
    return lang.startsWith('th') ? 'th' : 'en';
}
let currentLang = detectBrowserLanguage();

const I18N_UI = {
    th: {
        pageTitle: 'Camping มหาสนุก v1.1',
        homeAria: 'กลับสู่ชุมชน',
        langBtnText: '🌐 EN',
        langBtnAria: 'เปลี่ยนภาษาเป็นภาษาอังกฤษ',
        welcomeTitle: 'Camping มหาสนุก : Games4all 🏕️',
        welcomeSub1: 'ทอยลูกเต๋า ลุ้นทุกตา ฝ่าทุกเหตุการณ์จนถึงเส้นชัย',
        welcomeSub2: 'เกมนี้ออกแบบให้ผู้ใช้ Screen Reader และผู้เล่นทั่วไป สนุกร่วมกันได้อย่างเท่าเทียม',
        labelPlayerName: 'ชื่อผู้เล่นของคุณ:',
        placeholderPlayerName: 'กรอกชื่อของคุณที่นี่...',
        btnConfirmName: 'ยืนยันชื่อและเข้าสู่ล็อบบี้',
        lobbyHeading: 'ล็อบบี้การผจญภัย',
        btnCreateRoom: 'สร้างห้องใหม่ (Camping)',
        btnOpenManual: '📖 คู่มือการเล่นและกติกา',
        waitingRoomsTitle: 'ห้องที่กำลังรอผู้เล่น:',
        roomListAria: 'รายการห้องเกม',
        lobbyStatusWaiting: 'กำลังรอผู้เล่นเข้าร่วม (อย่างน้อย 2 คน)...',
        hostBotLabel: 'จัดการบอทในห้อง (รวมสูงสุด 6 คน):',
        btnRemoveBotAria: 'ลดจำนวนบอท',
        btnAddBotAria: 'เพิ่มจำนวนบอท',
        participantListLabel: 'รายชื่อผู้เล่นในห้อง',
        btnStartGame: '🚀 เริ่มเกมผจญภัย',
        manualPageTitle: 'คู่มือการเล่นและกติกา',
        btnCloseManual: 'ปิดคู่มือ',
        animStartTitle: 'การผจญภัยเริ่มต้นขึ้น!',
        animWinnerTitle: '🎉 ขอแสดงความยินดี! 🎉',
        boardAria: 'กระดานผจญภัย 80 ช่อง',
        btnRollDice: 'ทอยลูกเต๋า',
        resultTitle: 'จบการผจญภัย!',
        btnHomeResultText: 'กลับสู่ชุมชน',
        manualHtml: `
            <h3 style="color: var(--secondary); margin-bottom: 8px;">🏕️ เป้าหมายและกติกาพื้นฐาน</h3>
            <ul style="padding-left: 20px; margin-bottom: 15px;">
                <li><strong>เป้าหมาย:</strong> ทอยลูกเต๋าพาตัวละครเดินทางจากจุดเริ่มต้น (ช่อง 1) ไปยังเส้นชัยปราสาท (ช่อง 80) เป็นคนแรก</li>
                <li><strong>เวลา (Timer):</strong> มีเวลา 40 วินาทีในการตัดสินใจแต่ละเทิร์น หากหมดเวลาจะถูกข้ามอัตโนมัติ</li>
                <li><strong>ผู้เล่นหลุด / Host:</strong> หากผู้เล่นหลุดจะกลายเป็นบอทโดยอัตโนมัติ และหาก Host หลุด ห้องจะถูกปิดทันที</li>
            </ul>

            <h3 style="color: var(--secondary); margin-bottom: 8px;">🗺️ กลไกและประเภทช่องบนกระดาน</h3>
            <ul style="padding-left: 20px; margin-bottom: 15px;">
                <li><strong>จุดพักผ่อน (Rest):</strong> ช่องปลอดภัย ไม่มีเหตุการณ์พิเศษเกิดขึ้น</li>
                <li><strong>หีบสมบัติ & กล่องกุญแจ (Treasure/Key):</strong> ช่องเก็บสะสมกุญแจเพื่อใช้เปิดประตูทางลัด</li>
                <li><strong>ประตูทางลัด (Door):</strong> ใช้กุญแจเพื่อเปิดทางพุ่งไปยังช่องที่ไกลกว่าเดิม</li>
                <li><strong>วาร์ป & ลาภลอย (Forward/Bonus):</strong> ช่วยพัดพาตัวละครก้าวกระโดดไปข้างหน้า</li>
                <li><strong>น้ำเชี่ยว & ไซโคลน (Water/Warp):</strong> พัดตัวละครไปยังตำแหน่งที่กำหนด</li>
                <li><strong>หลุมพราง & ผีหลอก (Trap/Ghost):</strong> ช่องอันตรายที่อาจทำให้เสียตาเดินหรือถูกหน่วงเวลา</li>
            </ul>

            <h3 style="color: var(--secondary); margin-bottom: 8px;">⌨️ แป้นพิมพ์ลัด</h3>
            <ul style="padding-left: 20px; margin-bottom: 0;">
                <li><strong>Alt + r:</strong> กดทอยลูกเต๋า</li>
                <li><strong>Alt + a:</strong> อ่านสถานะตัวละครทุกคน</li>
            </ul>
        `
    },
    en: {
        pageTitle: 'Camping Adventure v1.1',
        homeAria: 'Back to Community',
        langBtnText: '🌐 TH',
        langBtnAria: 'Switch language to Thai',
        welcomeTitle: 'Camping Adventure : Games4all 🏕️',
        welcomeSub1: 'Roll the dice, face every event, and race to the finish!',
        welcomeSub2: 'Designed for Screen Reader users and sighted players to enjoy equally.',
        labelPlayerName: 'Your Player Name:',
        placeholderPlayerName: 'Enter your name here...',
        btnConfirmName: 'Confirm Name & Enter Lobby',
        lobbyHeading: 'Adventure Lobby',
        btnCreateRoom: 'Create New Room (Camping)',
        btnOpenManual: '📖 Game Manual & Rules',
        waitingRoomsTitle: 'Waiting Rooms:',
        roomListAria: 'Game room list',
        lobbyStatusWaiting: 'Waiting for players to join (at least 2)...',
        hostBotLabel: 'Manage Bots (Max 6 players total):',
        btnRemoveBotAria: 'Decrease bots',
        btnAddBotAria: 'Increase bots',
        participantListLabel: 'Players in Room',
        btnStartGame: '🚀 Start Adventure',
        manualPageTitle: 'Game Manual & Rules',
        btnCloseManual: 'Close Manual',
        animStartTitle: 'The Adventure Begins!',
        animWinnerTitle: '🎉 Congratulations! 🎉',
        boardAria: '80-Space Adventure Board',
        btnRollDice: 'Roll Dice',
        resultTitle: 'Adventure Complete!',
        btnHomeResultText: 'Back to Community',
        manualHtml: `
            <h3 style="color: var(--secondary); margin-bottom: 8px;">🏕️ Objective & Basic Rules</h3>
            <ul style="padding-left: 20px; margin-bottom: 15px;">
                <li><strong>Objective:</strong> Roll the dice to guide your character from Start (Space 1) to the Castle Finish (Space 80) first.</li>
                <li><strong>Timer:</strong> You have 40 seconds per turn. If time runs out, your turn will auto-play.</li>
                <li><strong>Disconnects / Host:</strong> Disconnected players automatically become bots. If the Host disconnects, the room closes immediately.</li>
            </ul>

            <h3 style="color: var(--secondary); margin-bottom: 8px;">🗺️ Board Mechanics & Space Types</h3>
            <ul style="padding-left: 20px; margin-bottom: 15px;">
                <li><strong>Rest:</strong> Safe space with no special event.</li>
                <li><strong>Treasure & Mystery Box (Treasure/Key):</strong> Collect Holy Armor or Ancient Keys to open shortcut doors.</li>
                <li><strong>Shortcut Door (Door):</strong> Use a key to warp ahead to a farther space.</li>
                <li><strong>Warp & Windfall (Forward/Bonus):</strong> Boosts your character forward or grants bonus items.</li>
                <li><strong>Rapids & Cyclone (Water/Warp):</strong> Sweeps your character forward or backward.</li>
                <li><strong>Trap & Ghost (Trap/Ghost):</strong> Dangerous spaces that push you backward unless blocked by Holy Armor.</li>
            </ul>

            <h3 style="color: var(--secondary); margin-bottom: 8px;">⌨️ Keyboard Shortcuts</h3>
            <ul style="padding-left: 20px; margin-bottom: 0;">
                <li><strong>Alt + r:</strong> Roll Dice</li>
                <li><strong>Alt + a:</strong> Announce all players' status</li>
            </ul>
        `
    }
};

function applyLanguageUI() {
    const t = I18N_UI[currentLang] || I18N_UI.th;
    document.documentElement.lang = currentLang;
    document.title = t.pageTitle;

    const setText = (id, text) => {
        const el = document.getElementById(id);
        if (el) el.textContent = text;
    };
    const setAria = (id, label) => {
        const el = document.getElementById(id);
        if (el) el.setAttribute('aria-label', label);
    };

    setAria('btn-home-welcome', t.homeAria);
    setAria('btn-home-lobby', t.homeAria);
    setAria('btn-home-result', t.homeAria);

    setText('btn-lang-welcome', t.langBtnText);
    setAria('btn-lang-welcome', t.langBtnAria);
    setText('btn-lang-lobby', t.langBtnText);
    setAria('btn-lang-lobby', t.langBtnAria);

    setText('welcome-title', t.welcomeTitle);
    setText('welcome-sub-1', t.welcomeSub1);
    setText('welcome-sub-2', t.welcomeSub2);
    setText('label-player-name', t.labelPlayerName);
    const nameInput = document.getElementById('player-name-input');
    if (nameInput) nameInput.placeholder = t.placeholderPlayerName;
    setText('btn-confirm-name-text', t.btnConfirmName);

    setText('lobby-heading', t.lobbyHeading);
    setText('btn-create-room-text', t.btnCreateRoom);
    setText('btn-open-manual-text', t.btnOpenManual);
    setText('waiting-rooms-title', t.waitingRoomsTitle);
    setAria('room-list', t.roomListAria);
    setText('lobby-status-text', t.lobbyStatusWaiting);
    setText('host-bot-label', t.hostBotLabel);
    setAria('btn-remove-bot', t.btnRemoveBotAria);
    setAria('btn-add-bot', t.btnAddBotAria);
    setText('participant-list-label', t.participantListLabel);
    setText('btn-start-game-text', t.btnStartGame);

    setText('manual-page-title', t.manualPageTitle);
    const manualSection = document.getElementById('manual-section');
    if (manualSection) manualSection.innerHTML = t.manualHtml;
    setText('btn-close-manual-text', t.btnCloseManual);

    setText('anim-start-title', t.animStartTitle);
    setText('anim-winner-title', t.animWinnerTitle);
    setAria('adventure-board', t.boardAria);
    setText('btn-roll-dice-text', t.btnRollDice);
    setText('result-title', t.resultTitle);
    setText('btn-home-result-text', t.btnHomeResultText);

    if (gameState) {
        updateLobbyUI();
    }
    if (lastRoomsSnapshot !== null && !myPlayerId) {
        renderRoomList(lastRoomsSnapshot);
    }
}

window.toggleLanguage = function() {
    playAudio('select.mp3');
    currentLang = (currentLang === 'th') ? 'en' : 'th';
    try {
        localStorage.setItem('adventure80_lang', currentLang);
    } catch (e) {}
    applyLanguageUI();
    if (currentLang === 'en') {
        announceSR('Language switched to English', 'polite');
    } else {
        announceSR('เปลี่ยนภาษาเป็นภาษาไทยแล้ว', 'polite');
    }
};

// Translate canonical Thai Firebase action messages to English only when currentLang === 'en'
function translateEventMessage(msg) {
    if (currentLang === 'th' || !msg) return msg;

    const spaceTypeMap = {
        'จุดพักผ่อน': 'Rest Space',
        'หีบสมบัติ': 'Treasure Chest',
        'วาร์ป': 'Warp',
        'หลุมพราง': 'Trap',
        'น้ำเชี่ยว': 'Rapids',
        'ผีหลอก': 'Ghost',
        'ไซโคลน': 'Cyclone',
        'กล่องลึกลับ': 'Mystery Box',
        'ประตูทางลัด': 'Shortcut Door',
        'ลาภลอย': 'Windfall',
        'เหตุการณ์ลับ': 'Secret Event',
        'พิเศษ': 'Special'
    };

    const restNameMap = {
        'แคมป์ไฟอบอุ่น': 'Warm Campfire',
        'โอเอซิสสงบเงียบ': 'Peaceful Oasis',
        'กระท่อมร้างกลางป่า': 'Forest Cabin',
        'ใต้ต้นไม้ใหญ่': 'Under the Great Tree',
        'ศาลาพักใจ': 'Resting Pavilion'
    };

    const doorNameMap = {
        'ประตูมิติ': 'Dimensional Door',
        'ประตูกาลเวลา': 'Time Door',
        'ประตูศิลาโบราณ': 'Ancient Stone Door',
        'ประตูเวทมนตร์': 'Magic Door',
        'ประตูลับ': 'Secret Door'
    };

    const secretNameMap = {
        'วิญญาณศักดิ์สิทธิ์': 'Holy Spirit',
        'สัตว์เวทย์นำทาง': 'Mystic Guide',
        'ภูตแห่งแสง': 'Light Fairy'
    };

    let m;
    if ((m = msg.match(/^(.+?)\s+ทอยลูกเต๋าได้\s+(\d+)$/))) {
        return `${m[1]} rolled a ${m[2]}`;
    }
    if ((m = msg.match(/^(.+?)\s+เดินจากช่อง\s+(\d+)\s+ไปยังช่อง\s+(\d+)$/))) {
        return `${m[1]} moved from space ${m[2]} to space ${m[3]}`;
    }
    if ((m = msg.match(/^(.+?)ตกช่อง\s+(\d+)\s+เป็นช่อง(.+)$/))) {
        const tName = spaceTypeMap[m[3]] || m[3];
        return `${m[1]} landed on space ${m[2]}, a ${tName}`;
    }
    if ((m = msg.match(/^(.+?)\s+มาถึงจุดพักผ่อน\s+(.+?)\s+ช่อง\s+(\d+)\s+ไม่มีเหตุการณ์พิเศษ จบการเดินทาง$/))) {
        const rName = restNameMap[m[2]] || m[2];
        return `${m[1]} arrived at ${rName} on space ${m[3]}. Safe rest, turn ends.`;
    }
    if (msg === 'พบหีบสมบัติ เปิดหีบ พบเกราะศักดิ์สิทธิ์ ได้รับเกราะศักดิ์สิทธิ์ 1 ชิ้น') {
        return 'Found a Treasure Chest! Opened it and received 1 Holy Armor.';
    }
    if (msg === 'พบหีบสมบัติ แต่หีบถูกเปิดไปแล้ว เหลือเพียงเศษฝุ่น ไม่ได้อะไรเลย') {
        return 'Found a Treasure Chest, but it was already empty. Only dust remains.';
    }
    if (msg === 'พบกล่องลึกลับ เปิดกล่อง พบกุญแจโบราณ ได้รับกุญแจโบราณ 1 ดอก') {
        return 'Found a Mystery Box! Opened it and received 1 Ancient Key.';
    }
    if (msg === 'พบกล่องลึกลับ แต่กล่องถูกเปิดไปแล้ว เหลือเพียงเศษฝุ่น ไม่ได้อะไรเลย') {
        return 'Found a Mystery Box, but it was already empty. Only dust remains.';
    }
    if ((m = msg.match(/^ใช้กุญแจเปิด(.+?)สำเร็จ วาร์ปไปยังช่อง\s+(\d+)$/))) {
        const dName = doorNameMap[m[1]] || m[1];
        return `Used a key to open the ${dName}! Warped to space ${m[2]}.`;
    }
    if (msg === 'เลือกที่จะไม่ใช้กุญแจ ประตูยังคงปิดอยู่') {
        return 'Chose not to use a key. The door remains closed.';
    }
    if ((m = msg.match(/^พบ(.+?)\s+แต่ไม่มีกุญแจ ไม่สามารถเปิดได้$/))) {
        const dName = doorNameMap[m[1]] || m[1];
        return `Found a ${dName}, but have no key to open it.`;
    }
    if (msg === 'ใช้เกราะศักดิ์สิทธิ์ป้องกันผลเสียสำเร็จ ปลอดภัยแล้ว!') {
        return 'Used Holy Armor to block the danger! Safe!';
    }
    if ((m = msg.match(/^เลือกไม่ใช้เกราะศักดิ์สิทธิ์ รับผลร้าย ถอยหลัง\s+(\d+)\s+ช่อง$/))) {
        return `Chose not to use Holy Armor. Moved back ${m[1]} spaces.`;
    }
    if ((m = msg.match(/^เกิดเหตุการณ์ร้าย ถอยหลัง\s+(\d+)\s+ช่อง$/))) {
        return `Bad event! Moved back ${m[1]} spaces.`;
    }
    if ((m = msg.match(/^โชคดี เดินหน้าเพิ่ม\s+(\d+)\s+ช่อง$/))) {
        return `Lucky! Moved forward ${m[1]} spaces.`;
    }
    if ((m = msg.match(/^พลังพิเศษ เดินหน้า\s+(\d+)\s+ช่อง$/))) {
        return `Special boost! Moved forward ${m[1]} spaces.`;
    }
    if (msg === 'ลาภลอย! ได้รับกุญแจ 1 ชิ้น') {
        return 'Windfall! Received 1 Key.';
    }
    if (msg === 'ลาภลอย! ได้รับเกราะศักดิ์สิทธิ์ 1 ชิ้น') {
        return 'Windfall! Received 1 Holy Armor.';
    }
    if ((m = msg.match(/^ลาภลอย! เกิดการวาร์ป เดินหน้า\s+(\d+)\s+ช่อง$/))) {
        return `Windfall! Warped forward ${m[1]} spaces.`;
    }
    if ((m = msg.match(/^พบเหตุการณ์ลับ:\s+(.+?)\s+พาเดินหน้า\s+(\d+)\s+ช่อง$/))) {
        const sName = secretNameMap[m[1]] || m[1];
        return `Secret Event: ${sName} carried you forward ${m[2]} spaces.`;
    }
    if (msg === 'กลับมายังจุดเริ่มต้น') {
        return 'Returned to the starting space.';
    }
    if ((m = msg.match(/^ตรวจสอบช่องใหม่: กลับมาช่อง\s+(\d+)\s+ที่เคยผ่านมาแล้ว หยุดการเดินทางเพื่อป้องกันการวนลูป$/))) {
        return `Checking new space: returned to visited space ${m[1]}. Stopping to prevent a loop.`;
    }
    if (msg === 'ตรวจสอบช่องใหม่: ครบขีดจำกัดเหตุการณ์ 3 ครั้งแล้ว จบการเคลื่อนที่สำหรับเทิร์นนี้ทันที') {
        return 'Checking new space: reached the 3-event chain limit. Turn ends.';
    }
    if ((m = msg.match(/^ตรวจสอบช่องใหม่: ไปยังช่อง\s+(\d+)$/))) {
        return `Checking new space: moved to space ${m[1]}.`;
    }

    return msg;
}

// Animal Avatars Pool (16 Cute Animals)
const ANIMAL_POOL = [
    { icon: '🦊', name: 'จิ้งจอก' }, { icon: '🐼', name: 'แพนด้า' },
    { icon: '🐰', name: 'กระต่าย' }, { icon: '🐨', name: 'โคอาลา' },
    { icon: '🐯', name: 'เสือ' },     { icon: '🦁', name: 'สิงโต' },
    { icon: '🐸', name: 'กบ' },     { icon: '🐵', name: 'ลิง' },
    { icon: '🐧', name: 'เพนกวิน' }, { icon: '🐻', name: 'หมี' },
    { icon: '🦝', name: 'แร็กคูน' }, { icon: '🦄', name: 'ยูนิคอร์น' },
    { icon: '🐹', name: 'แฮมสเตอร์' }, { icon: '🐱', name: 'แมว' },
    { icon: '🐶', name: 'สุนัข' },   { icon: '🐾', name: 'แพนด้าแดง' }
];

// Global Game States
let myPlayerName = '';
let myPlayerId = null;
let currentRoomId = null;
let isHost = false;
let gameState = null;
let roomListener = null;
let roomListListener = null;
let lastRoomsSnapshot = null;
let localLastActionTs = 0;
let speechQueue = [];
let isSpeaking = false;
let lastAnnouncedTurnKey = null;
let previousPlayersState = {};
let myTurnTimer = null;
let isStartingGame = false;

// Audio System States
let bgmSource = null;
let isBgmActive = false;
let audioBuffers = {};
let audioLoadingPromises = {};
let previousHumanCount = null;
let previousBotCount = null;
let isShowingWinnerScene = false;
let audioUnlocked = false;
let unlockPromise = null;
let hasPreloadedAudio = false;
let isConfirmingName = false;

// All Audio Files Used in Game (Preserving exact file names)
const AUDIO_FILES_LIST = [
    'select.mp3', '1.mp3', 'start.mp3', 'bgm.mp3', 'win.mp3',
    'turn.mp3', 'abc.mp3', 'dice.mp3', 'walk.mp3', 'sabuy.mp3',
    'box1.mp3', 'forward.mp3', 'hole.mp3', 'water.mp3', 'ghost.mp3',
    'cyclone.mp3', 'box2.mp3', 'door1.mp3', 'box4.mp3', 'box3.mp3',
    'akey3.mp3', 'akey2.mp3', 'akey1.mp3', 'door2.mp3', 'shieldno.mp3',
    'shield.mp3'
];

// Audio Context Web API (Web Audio API) - Single Instance Management
const AudioContextClass = window.AudioContext || window.webkitAudioContext;
let audioCtx = AudioContextClass ? new AudioContextClass() : null;

function getAudioContext() {
    if (!audioCtx && AudioContextClass) {
        try {
            audioCtx = new AudioContextClass();
        } catch (e) {
            console.error('AudioContext creation error:', e);
        }
    }
    return audioCtx;
}

// Play 1-frame silent buffer inside User Gesture to unlock iOS/Safari WebKit audio pipeline
function playSilentUnlockBuffer(ctx) {
    if (!ctx) return;
    try {
        const buffer = ctx.createBuffer(1, 1, 22050);
        const source = ctx.createBufferSource();
        source.buffer = buffer;
        source.connect(ctx.destination);
        if (typeof source.start === 'function') {
            source.start(0);
        } else if (typeof source.noteOn === 'function') {
            source.noteOn(0);
        }
    } catch (e) {
        // Ignore silent buffer errors on strict states
    }
}

// Unlock & Resume AudioContext safely with Promise deduplication
function unlockAudioContext() {
    const ctx = getAudioContext();
    if (!ctx) return Promise.resolve(false);

    // If already running and unlocked, return immediately without redundant calls
    if (ctx.state === 'running' && audioUnlocked) {
        return Promise.resolve(true);
    }

    // Prevent concurrent unlock calls
    if (unlockPromise) {
        return unlockPromise;
    }

    // Trigger silent buffer synchronously within the user gesture stack frame
    playSilentUnlockBuffer(ctx);

    unlockPromise = (async () => {
        try {
            if (ctx.state === 'suspended' || ctx.state === 'interrupted') {
                await ctx.resume();
            }
            if (ctx.state === 'running') {
                playSilentUnlockBuffer(ctx);
                audioUnlocked = true;
                return true;
            }
            return false;
        } catch (err) {
            console.warn('AudioContext resume warning:', err);
            return ctx.state === 'running';
        } finally {
            unlockPromise = null;
        }
    })();

    return unlockPromise;
}

// Handler for User Gesture Audio Recovery (Initial Unlock & Background/Lock Recovery)
const handleUserGestureAudio = () => {
    const ctx = getAudioContext();
    if (!ctx) return;
    if (!audioUnlocked || ctx.state === 'suspended' || ctx.state === 'interrupted') {
        unlockAudioContext().then((ready) => {
            if (ready && isBgmActive && !bgmSource) {
                const gameScreen = document.getElementById('screen-game');
                if (gameScreen && gameScreen.classList.contains('active-screen') && gameState && gameState.status === 'playing') {
                    playAudio('bgm.mp3', true);
                }
            }
        });
    }
};

document.addEventListener('click', handleUserGestureAudio, { capture: true, passive: true });
document.addEventListener('touchstart', handleUserGestureAudio, { capture: true, passive: true });
document.addEventListener('touchend', handleUserGestureAudio, { capture: true, passive: true });
document.addEventListener('pointerdown', handleUserGestureAudio, { capture: true, passive: true });
document.addEventListener('keydown', handleUserGestureAudio, { capture: true, passive: true });

// Monitor AudioContext state changes and Page Visibility for Safari/iOS app switching & screen lock
if (audioCtx && typeof audioCtx.addEventListener === 'function') {
    audioCtx.addEventListener('statechange', () => {
        if (audioCtx.state === 'running') {
            audioUnlocked = true;
        }
    });
}

document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') {
        const ctx = getAudioContext();
        if (ctx && (ctx.state === 'suspended' || ctx.state === 'interrupted')) {
            ctx.resume().catch(() => {});
        }
    }
});

window.addEventListener('pageshow', () => {
    const ctx = getAudioContext();
    if (ctx && (ctx.state === 'suspended' || ctx.state === 'interrupted')) {
        ctx.resume().catch(() => {});
    }
});

window.addEventListener('focus', () => {
    const ctx = getAudioContext();
    if (ctx && (ctx.state === 'suspended' || ctx.state === 'interrupted')) {
        ctx.resume().catch(() => {});
    }
});

// Cross-browser & Legacy Safari compatible decodeAudioData wrapper
function decodeAudioDataSafe(ctx, arrayBuffer) {
    return new Promise((resolve, reject) => {
        try {
            const bufferCopy = arrayBuffer.slice(0);
            const decodeResult = ctx.decodeAudioData(
                bufferCopy,
                (decodedBuffer) => resolve(decodedBuffer),
                (err) => reject(err)
            );
            if (decodeResult && typeof decodeResult.then === 'function') {
                decodeResult.then(resolve).catch(reject);
            }
        } catch (e) {
            reject(e);
        }
    });
}

async function loadAudio(name) {
    if (!name) return null;
    if (audioBuffers[name]) return audioBuffers[name];
    if (audioLoadingPromises[name]) return audioLoadingPromises[name];

    const ctx = getAudioContext();
    if (!ctx) return null;

    audioLoadingPromises[name] = (async () => {
        try {
            const response = await fetch(`audio/${name}`);
            if (!response.ok) {
                throw new Error(`HTTP ${response.status} while loading audio/${name}`);
            }
            const arrayBuffer = await response.arrayBuffer();
            const audioBuffer = await decodeAudioDataSafe(ctx, arrayBuffer);
            if (audioBuffer) {
                audioBuffers[name] = audioBuffer;
            }
            return audioBuffer;
        } catch (e) {
            console.error('Audio load error:', name, e);
            return null;
        } finally {
            delete audioLoadingPromises[name];
        }
    })();

    return audioLoadingPromises[name];
}

// Non-blocking background audio preloader triggered after player confirms name
function preloadAllAudio() {
    if (hasPreloadedAudio) return;
    hasPreloadedAudio = true;

    const priorityFiles = ['select.mp3', '1.mp3', 'start.mp3', 'turn.mp3', 'abc.mp3', 'dice.mp3', 'walk.mp3'];
    priorityFiles.forEach((file) => {
        loadAudio(file);
    });

    setTimeout(() => {
        AUDIO_FILES_LIST.forEach((file) => {
            if (!audioBuffers[file] && !audioLoadingPromises[file]) {
                loadAudio(file);
            }
        });
    }, 300);
}

// Ensure AudioContext is ready before playing without hanging non-gesture callbacks (Firebase/Timers)
async function ensureAudioContextReady() {
    const ctx = getAudioContext();
    if (!ctx) return false;
    if (ctx.state === 'running') return true;

    if (ctx.state === 'suspended' || ctx.state === 'interrupted') {
        try {
            await Promise.race([
                unlockAudioContext(),
                new Promise((resolve) => setTimeout(resolve, 250))
            ]);
        } catch (e) {
            // Ignore resume rejections outside of user gesture
        }
    }
    return ctx.state === 'running';
}

function playAudio(name, loop = false) {
    return new Promise(async (resolve) => {
        try {
            const ctx = getAudioContext();
            if (!ctx) {
                resolve();
                return;
            }

            if (loop && name === 'bgm.mp3') {
                isBgmActive = true;
            }

            const [buffer] = await Promise.all([
                loadAudio(name),
                ensureAudioContextReady()
            ]);

            if (!buffer) {
                resolve();
                return;
            }

            if (ctx.state === 'suspended' || ctx.state === 'interrupted') {
                try {
                    await Promise.race([
                        ctx.resume(),
                        new Promise((r) => setTimeout(r, 100))
                    ]);
                } catch (e) {}
            }

            if (ctx.state !== 'running') {
                resolve();
                return;
            }

            const source = ctx.createBufferSource();
            source.buffer = buffer;
            source.connect(ctx.destination);
            source.loop = loop;

            let settled = false;
            const finish = () => {
                if (!settled) {
                    settled = true;
                    resolve(source);
                }
            };

            source.onended = () => {
                if (loop && name === 'bgm.mp3' && bgmSource === source) {
                    bgmSource = null;
                }
                finish();
            };

            if (loop && name === 'bgm.mp3') {
                if (bgmSource) {
                    try {
                        bgmSource.onended = null;
                        bgmSource.stop();
                        bgmSource.disconnect();
                    } catch (e) {}
                }
                bgmSource = source;
            }

            source.start(0);

            if (loop) {
                finish();
            }
        } catch (err) {
            console.warn('Audio playback error:', name, err);
            resolve();
        }
    });
}

async function playAudioSequence(list) {
    for (const name of list) {
        await playAudio(name);
    }
}

function stopBGM() {
    isBgmActive = false;
    if (bgmSource) {
        try {
            bgmSource.onended = null;
            bgmSource.stop();
            bgmSource.disconnect();
        } catch (e) {}
        bgmSource = null;
    }
}

// ARIA Queue Announcement Manager (Local sequence handler)
function announceSR(text, priority = 'polite') {
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
            // ให้ระยะเวลาการพูดตามความยาวของข้อความคร่าวๆ (อย่างน้อย 1500ms ป้องกันการทับกัน)
            const duration = Math.max(1500, item.text.length * 50);
            setTimeout(() => {
                targetEl.textContent = '';
                isSpeaking = false;
                processSpeechQueue();
            }, duration);
        }, 100);
    } else {
        isSpeaking = false;
    }
}

// Helper: Delay Async execution
function delayAsync(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Focus Management Screen Switcher
function switchScreen(screenId, focusHeadingId = null) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active-screen'));
    const target = document.getElementById(screenId);
    if (target) {
        target.classList.add('active-screen');
        if (focusHeadingId) {
            const h = document.getElementById(focusHeadingId);
            if (h) {
                h.setAttribute('tabindex', '-1');
                h.focus();
                // ป้องกัน iOS VoiceOver ดึงโฟกัสกลับมาอ่าน Heading ซ้ำเมื่อมีการเปลี่ยนสถานะโฟกัสในหน้าจอ
                const cleanupFocus = () => {
                    h.removeAttribute('tabindex');
                    h.removeEventListener('blur', cleanupFocus);
                };
                h.addEventListener('blur', cleanupFocus);
            }
        }
    }
    
    // จัดการซ่อนปุ่มกลับสู่ชุมชน/หน้าแรก ในระหว่างการเล่นเกม
    const homeBtns = document.querySelectorAll('.btn-home, #btn-home');
    const communityBtns = document.querySelectorAll('.btn-community, #btn-community');
    
    homeBtns.forEach(btn => {
        btn.style.display = (screenId === 'screen-welcome') ? 'block' : 'none';
    });
    communityBtns.forEach(btn => {
        btn.style.display = (screenId === 'screen-result') ? 'block' : 'none';
    });
}

// Name and Navigation (Primary Audio Unlock & Preload Entry Point)
window.confirmNameAndEnterLobby = function() {
    if (isConfirmingName) return;
    const welcomeScreen = document.getElementById('screen-welcome');
    if (welcomeScreen && !welcomeScreen.classList.contains('active-screen')) return;

    isConfirmingName = true;
    setTimeout(() => {
        isConfirmingName = false;
    }, 400);

    // Unlock AudioContext immediately inside direct User Gesture (Button click / Virtual Enter / Physical Enter)
    unlockAudioContext();
    preloadAllAudio();

    playAudio('select.mp3');
    const nameInput = document.getElementById('player-name-input');
    myPlayerName = nameInput.value.trim() || (currentLang === 'en' ? 'New Player' : 'ผู้เล่นใหม่');
    switchScreen('screen-lobby', 'lobby-heading');
    if (currentLang === 'en') {
        announceSR(`Welcome ${myPlayerName} to the Adventure Lobby`);
    } else {
        announceSR(`ยินดีต้อนรับ ${myPlayerName} เข้าสู่ล็อบบี้การผจญภัย`);
    }
};

window.leaveRoomAndGoHome = function() {
    if (currentRoomId) window.leaveRoom();
    switchScreen('screen-welcome', 'welcome-title');
};

window.toggleManual = function() {
    playAudio('select.mp3');
    const m = document.getElementById('manual-section');
    if (m.style.display === 'none') {
        m.style.display = 'block';
        m.focus();
        announceSR(currentLang === 'en' ? "Game manual opened" : "เปิดคู่มือการเล่นแล้ว");
    } else {
        m.style.display = 'none';
        announceSR(currentLang === 'en' ? "Game manual closed" : "ปิดคู่มือการเล่นแล้ว");
    }
};

// Render Room List helper
function renderRoomList(rooms) {
    const container = document.getElementById('room-list');
    if (!container || myPlayerId) return;

    container.innerHTML = '';
    let count = 0;
    for (const [rId, rData] of Object.entries(rooms || {})) {
        if (rData.status === 'waiting') {
            count++;
            const item = document.createElement('div');
            item.className = 'room-item';
            const pCount = rData.players ? Object.keys(rData.players).length : 0;
            if (currentLang === 'en') {
                item.innerHTML = `<span><strong>Room ${rId}</strong> (${pCount}/6 players)</span>`;
            } else {
                item.innerHTML = `<span><strong>ห้อง ${rId}</strong> (${pCount}/6 คน)</span>`;
            }
            item.onclick = () => window.joinAdventureRoom(rId);
            container.appendChild(item);
        }
    }
    if (count === 0) {
        const emptyText = currentLang === 'en'
            ? 'No rooms waiting for players...'
            : 'ไม่มีห้องที่กำลังรอผู้เล่นอยู่...';
        container.innerHTML = `<div style="color:#a0a5b5; padding:10px;">${emptyText}</div>`;
    }
}

// Initialize Room Listener for Main Menu
function initRoomListListener() {
    const roomsRef = ref(db, 'games/Adventure80/rooms');
    roomListListener = onValue(roomsRef, (snapshot) => {
        const rooms = snapshot.val() || {};
        lastRoomsSnapshot = rooms;
        renderRoomList(rooms);
    });
}

// Board Random Generator (Host Authoritative - with High Density Requirements)
function generateBoardConfiguration() {
    const available = [];
    for (let i = 2; i <= 79; i++) available.push(i);

    // Shuffle helper
    function shuffle(arr) {
        for (let i = arr.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [arr[i], arr[j]] = [arr[j], arr[i]];
        }
    }
    shuffle(available);

    const keys = [];
    const doors = [];
    
    // Select 10 Keys and 12 Doors (More doors and keys as per prompt)
    for (let k = 0; k < 10; k++) {
        let keySpace = available.pop();
        while (keySpace >= 70 && available.length > 0) {
            available.unshift(keySpace);
            keySpace = available.pop();
        }

        let possibleDoorSpaces = available.filter(s => s > 10 && (s - keySpace) >= 6);
        if (possibleDoorSpaces.length === 0) {
            keySpace = 12 + k * 2; // fallback
            possibleDoorSpaces = [keySpace + 7];
        }
        const doorSpace = possibleDoorSpaces[Math.floor(Math.random() * possibleDoorSpaces.length)];
        
        const idx = available.indexOf(doorSpace);
        if (idx !== -1) available.splice(idx, 1);

        const dest = Math.min(79, doorSpace + Math.floor(Math.random() * 14) + 5);
        keys.push(keySpace);
        doors.push({ space: doorSpace, dest });
    }
    
    // Add 2 extra standalone doors just to increase density
    for(let k = 0; k < 2; k++) {
        if (available.length > 0) {
            const extraDoor = available.pop();
            const dest = Math.min(79, extraDoor + Math.floor(Math.random() * 14) + 5);
            doors.push({ space: extraDoor, dest });
        }
    }

    const layout = {};
    keys.forEach(k => layout[k] = { type: 'key', charges: 2 });
    doors.forEach(d => layout[d.space] = { type: 'door', dest: d.dest });

    const specialTypes = [
        { type: 'bonus', count: 3 },
        { type: 'secret', count: 3 },
        { type: 'treasure', count: 10 },
        { type: 'forward', count: 8 },  
        { type: 'trap', count: 9 },     
        { type: 'water', count: 9 },    
        { type: 'ghost', count: 7 },    
        { type: 'warp', count: 4 },     
        { type: 'rest', count: 12 }     
    ];

    specialTypes.forEach(st => {
        for (let c = 0; c < st.count; c++) {
            if (available.length > 0) {
                const sp = available.pop();
                if (st.type === 'treasure') {
                    layout[sp] = { type: st.type, charges: 2 };
                } else {
                    layout[sp] = { type: st.type };
                }
            }
        }
    });

    return layout;
}

// Create Room Handler
window.createAdventureRoom = function() {
    if (!myPlayerName) myPlayerName = currentLang === 'en' ? 'Player 1' : 'ผู้เล่น 1';

    playAudio('1.mp3');
    isStartingGame = false;

    const counterRef = ref(db, 'games/Adventure80/room_counter');
    runTransaction(counterRef, (cur) => (cur || 0) + 1).then((res) => {
        if (res.committed) {
            const count = res.snapshot.val();
            currentRoomId = 'Adventure' + String(count).padStart(5, '0');
            myPlayerId = 'p1';
            isHost = true;

            const roomRef = ref(db, `games/Adventure80/rooms/${currentRoomId}`);
            
            // Host disconnecting closes the room entirely
            onDisconnect(roomRef).remove();

            const initialPlayers = {
                p1: { name: myPlayerName, animal: ANIMAL_POOL[0], isBot: false, pos: 1, keys: 0, armor: 0 }
            };

            set(roomRef, {
                status: 'waiting',
                players: initialPlayers,
                botCount: 0,
                turnIndex: 0,
                boardConfig: null,
                lastAction: { msg: '', ts: 0 }
            });

            document.getElementById('lobby-menu-section').style.display = 'none';
            document.getElementById('lobby-room-section').style.display = 'block';
            setupRoomListener();
            if (currentLang === 'en') {
                announceSR(`Room created successfully. Room ID is ${currentRoomId}`);
            } else {
                announceSR(`สร้างห้องสำเร็จ รหัสห้องคือ ${currentRoomId}`);
            }
        }
    });
};

// Join Room Handler
window.joinAdventureRoom = function(rId) {
    if (!myPlayerName) myPlayerName = currentLang === 'en' ? 'New Player' : 'ผู้เล่นเข้าใหม่';

    currentRoomId = rId;
    isHost = false;

    const roomRef = ref(db, `games/Adventure80/rooms/${currentRoomId}`);
    runTransaction(roomRef, (roomData) => {
        if (!roomData || roomData.status !== 'waiting') return;
        const pKeys = Object.keys(roomData.players || {});
        if (pKeys.length >= 6) return;

        let nextSlot = 'p1';
        for (let i = 1; i <= 6; i++) {
            if (!pKeys.includes('p' + i)) {
                nextSlot = 'p' + i;
                break;
            }
        }

        const usedAnimals = Object.values(roomData.players).map(p => p.animal.icon);
        const availAnimals = ANIMAL_POOL.filter(a => !usedAnimals.includes(a.icon));
        const assignedAnimal = availAnimals[0] || ANIMAL_POOL[0];

        if (!roomData.players) roomData.players = {};
        roomData.players[nextSlot] = {
            name: myPlayerName,
            animal: assignedAnimal,
            isBot: false,
            pos: 1,
            keys: 0,
            armor: 0
        };

        myPlayerId = nextSlot;
        return roomData;
    }).then((res) => {
        if (res.committed) {
            playAudio('select.mp3');
            // Client disconnecting turns them into a bot
            onDisconnect(ref(db, `games/Adventure80/rooms/${currentRoomId}/players/${myPlayerId}`)).update({
                isBot: true,
                name: "บอท" + myPlayerName
            });
            
            document.getElementById('lobby-menu-section').style.display = 'none';
            document.getElementById('lobby-room-section').style.display = 'block';
            
            setupRoomListener();
            if (currentLang === 'en') {
                announceSR(`Joined room ${currentRoomId} successfully`);
            } else {
                announceSR(`เข้าร่วมห้อง ${currentRoomId} เรียบร้อยแล้ว`);
            }
            document.getElementById('lobby-room-title').focus();
        }
    });
};

// Adjust Bot Count (Host Only)
window.adjustBot = function(delta) {
    if (!isHost || !gameState) return;
    const curBots = gameState.botCount || 0;
    const humanCount = Object.values(gameState.players).filter(p => !p.isBot).length;
    const newBots = curBots + delta;

    if (newBots < 0 || (humanCount + newBots) > 6) return;

    const roomRef = ref(db, `games/Adventure80/rooms/${currentRoomId}`);
    const updatedPlayers = { ...gameState.players };

    if (delta > 0) {
        const botSlot = 'p' + (humanCount + newBots);
        const botNames = ['บอทจอมทัพ', 'บอทน้ำพุ', 'บอทสายรุ้ง', 'บอทวายุ', 'บอทปัญญา'];
        const usedAnimals = Object.values(updatedPlayers).map(p => p.animal.icon);
        const availAnimals = ANIMAL_POOL.filter(a => !usedAnimals.includes(a.icon));

        updatedPlayers[botSlot] = {
            name: botNames[newBots - 1] || `บอท ${newBots}`,
            animal: availAnimals[0] || ANIMAL_POOL[1],
            isBot: true,
            pos: 1,
            keys: 0,
            armor: 0
        };
    } else {
        const botSlots = Object.keys(updatedPlayers).filter(k => updatedPlayers[k].isBot);
        if (botSlots.length > 0) {
            delete updatedPlayers[botSlots[botSlots.length - 1]];
        }
    }

    update(roomRef, {
        botCount: newBots,
        players: updatedPlayers
    });
};

// Listen to Room Realtime Data
function setupRoomListener() {
    const roomRef = ref(db, `games/Adventure80/rooms/${currentRoomId}`);
    roomListener = onValue(roomRef, (snapshot) => {
        gameState = snapshot.val();
        if (!gameState) {
            announceSR(
                currentLang === 'en' ? "The host left the game. Room closed." : "เจ้าของห้องออกจากเกม ห้องถูกปิด",
                'assertive'
            );
            setTimeout(() => {
                window.location.href = '../index.html';
            }, 3000);
            return;
        }

        // Check for client joined or bot changed for select.mp3
        if (gameState.status === 'waiting' && gameState.players) {
            let currentHumanCount = Object.values(gameState.players).filter(p => !p.isBot).length;
            let currentBotCount = gameState.botCount || 0;

            if (previousHumanCount !== null && currentHumanCount > previousHumanCount) {
                playAudio('select.mp3');
            }
            if (previousBotCount !== null && currentBotCount !== previousBotCount) {
                playAudio('select.mp3');
            }

            previousHumanCount = currentHumanCount;
            previousBotCount = currentBotCount;
        }

        // Check for disconnected players converting to bots
        if (gameState.players) {
            for (const pId in gameState.players) {
                if (previousPlayersState[pId] && !previousPlayersState[pId].isBot && gameState.players[pId].isBot) {
                    if (currentLang === 'en') {
                        announceSR(`${previousPlayersState[pId].name} disconnected and was converted to a bot`, 'polite');
                    } else {
                        announceSR(`${previousPlayersState[pId].name} หลุดจากห้อง แปลงเป็นบอท`, 'polite');
                    }
                }
            }
            previousPlayersState = JSON.parse(JSON.stringify(gameState.players));
        }

        updateLobbyUI();

        // Synced Actions Announcements Listener & Audio sync
        if (gameState.lastAction && gameState.lastAction.ts > localLastActionTs) {
            localLastActionTs = gameState.lastAction.ts;
            announceSR(translateEventMessage(gameState.lastAction.msg), 'polite');
            if (gameState.lastAction.audioKeys && gameState.lastAction.audioKeys.length > 0) {
                playAudioSequence(gameState.lastAction.audioKeys);
            }
            
            // --- SYNC DICE VISUAL: อัปเดตตัวเลขผลลูกเต๋าให้ทุกอุปกรณ์เห็นตรงกันจากการทอยจริง ---
            if (gameState.lastAction.msg.includes('ทอยลูกเต๋าได้')) {
                const diceMatch = gameState.lastAction.msg.match(/ทอยลูกเต๋าได้\s*(\d+)/);
                if (diceMatch) {
                    const diceEl = document.getElementById('dice-visual');
                    if (diceEl) {
                        diceEl.textContent = diceMatch[1];
                    }
                }
            }
            // -------------------------------------------------------------------------

            // --- VISUAL LAYER: Trigger Event Effect ---
            const currentTurnKey = gameState.playerOrder ? gameState.playerOrder[gameState.turnIndex] : null;
            const pData = (currentTurnKey && gameState.players) ? gameState.players[currentTurnKey] : null;
            VisualController.triggerFromMessage(gameState.lastAction.msg, pData);
            // ------------------------------------------
        }

        if (gameState.status === 'playing') {
            if (document.getElementById('screen-game').classList.contains('active-screen') === false) {
                showStartGameAnimation();
            } else {
                updateGameUI();
            }
        } else if (gameState.status === 'ended') {
            if (!isShowingWinnerScene && !document.getElementById('screen-result').classList.contains('active-screen')) {
                showWinnerAnimationAndResult();
            }
        }
    });
}

function showStartGameAnimation() {
    stopBGM();
    const overlay = document.getElementById('anim-start-overlay');
    overlay.style.display = 'flex';
    announceSR(
        currentLang === 'en' ? 'The adventure has begun! Get ready!' : 'การผจญภัยเริ่มต้นขึ้นแล้ว เตรียมตัวให้พร้อม!',
        'assertive'
    );
    
    // --- VISUAL LAYER: START EFFECTS ---
    VisualController.playStartEffects(overlay);
    // -----------------------------------
    
    playAudio('start.mp3').then(() => {
        // Wait 0.2s before playing BGM as requested
        setTimeout(() => {
            playAudio('bgm.mp3', true);
        }, 200);
    });

    // Start Animation lasts 4.89s (aligned with start.mp3 length)
    setTimeout(() => {
        overlay.style.display = 'none';
        switchScreen('screen-game', 'game-status-bar');
        renderBoardGrid();
        updateGameUI();
    }, 4890);
}

function showWinnerAnimationAndResult() {
    isShowingWinnerScene = true;
    stopBGM();

    const overlay = document.getElementById('anim-winner-overlay');
    const winnerId = gameState.winnerId;
    const winner = gameState.players[winnerId];

    document.getElementById('winner-anim-character').textContent = winner.animal.icon;
    document.getElementById('winner-anim-name').textContent = winner.name;
    
    overlay.style.display = 'flex';
    overlay.classList.add('active-winner-anim');

    // --- VISUAL LAYER: WINNER EFFECTS ---
    VisualController.playWinnerEffects(overlay);
    // ------------------------------------

    setTimeout(() => {
        playAudio('win.mp3');
    }, 100);

    // After 5s animation finishes, transition to Result screen
    setTimeout(() => {
        overlay.classList.remove('active-winner-anim');
        overlay.style.display = 'none';
        showResultScreen();
    }, 5000);
}

// Update Lobby Interface
function updateLobbyUI() {
    if (!gameState) return;
    document.getElementById('lobby-room-title').textContent = currentLang === 'en'
        ? `Game Room: ${currentRoomId}`
        : `ห้องเกม: ${currentRoomId}`;
    
    const pList = Object.values(gameState.players || {});
    document.getElementById('participant-count').textContent = `${pList.length}/6`;

    const listContainer = document.getElementById('participant-list');
    listContainer.innerHTML = '';
    const botTag = currentLang === 'en' ? '(Bot)' : '(บอท)';
    pList.forEach(p => {
        const card = document.createElement('div');
        card.className = 'participant-card';
        card.innerHTML = `<span class="participant-avatar">${p.animal.icon}</span> <span>${p.name} ${p.isBot ? botTag : ''}</span>`;
        listContainer.appendChild(card);
    });

    const hostPanel = document.getElementById('host-controls');
    const startBtn = document.getElementById('btn-start-game');

    if (isHost) {
        hostPanel.style.display = 'block';
        startBtn.style.display = 'block';
        startBtn.disabled = isStartingGame || gameState.status !== 'waiting' || pList.length < 2;
        document.getElementById('bot-count-display').textContent = gameState.botCount || 0;
        
        // จัดการสถานะปุ่มเพิ่ม/ลดบอท
        const btnAddBot = document.querySelector('button[onclick*="adjustBot(1)"]');
        const btnRemoveBot = document.querySelector('button[onclick*="adjustBot(-1)"]');
        
        if (btnAddBot) {
            const humanCount = pList.filter(p => !p.isBot).length;
            const botCount = gameState.botCount || 0;
            btnAddBot.disabled = (humanCount + botCount >= 6);
        }
        if (btnRemoveBot) {
            btnRemoveBot.disabled = ((gameState.botCount || 0) === 0);
        }
    } else {
        hostPanel.style.display = 'none';
        startBtn.style.display = 'none';
    }
}

// Start Game Handler (Host Only)
window.startAdventureGame = function() {
    if (!isHost || isStartingGame || !gameState || gameState.status !== 'waiting') return;

    isStartingGame = true;
    const startBtn = document.getElementById('btn-start-game');
    if (startBtn) startBtn.disabled = true;

    lastAnnouncedTurnKey = null;

    try {
        const boardConfig = generateBoardConfiguration();
        const roomRef = ref(db, `games/Adventure80/rooms/${currentRoomId}`);

        const pKeys = Object.keys(gameState.players || {});
        if (pKeys.length < 2) {
            isStartingGame = false;
            if (startBtn) startBtn.disabled = false;
            return;
        }

        for (let i = pKeys.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [pKeys[i], pKeys[j]] = [pKeys[j], pKeys[i]];
        }

        update(roomRef, {
            status: 'playing',
            boardConfig,
            turnIndex: 0,
            playerOrder: pKeys
        }).catch((err) => {
            console.error('Error starting game:', err);
            isStartingGame = false;
            if (startBtn) startBtn.disabled = false;
        });
    } catch (err) {
        console.error('Error starting game:', err);
        isStartingGame = false;
        if (startBtn) startBtn.disabled = false;
    }
};

// Render Serpentine Board Grid 8x10 with Accessible ARIA Names
function renderBoardGrid() {
    const grid = document.getElementById('adventure-board');
    grid.setAttribute('aria-hidden', 'true');
    grid.innerHTML = '';
    const config = gameState.boardConfig || {};

    for (let row = 0; row < 8; row++) {
        for (let col = 0; col < 10; col++) {
            let num;
            if (row % 2 === 0) {
                num = row * 10 + col + 1;
            } else {
                num = row * 10 + (10 - col);
            }

            const cell = document.createElement('div');
            cell.className = 'cell';
            cell.id = `cell-${num}`;
            cell.setAttribute('role', 'gridcell');

            let iconHtml = '';
            let accessibleName = currentLang === 'en' ? `Space ${num}` : `ช่อง ${num}`;

            if (num === 1) {
                cell.classList.add('cell-start');
                iconHtml = '🏕️';
                accessibleName += currentLang === 'en' ? ' Start' : ' จุดเริ่มต้น';
            } else if (num === 80) {
                cell.classList.add('cell-finish');
                iconHtml = '🏰';
                accessibleName += currentLang === 'en' ? ' Castle Finish' : ' เส้นชัยปราสาท';
            } else if (config[num]) {
                const sp = config[num];
                cell.classList.add(`cell-${sp.type}`);
                
                if (sp.type === 'rest') { iconHtml = '⛺'; accessibleName += currentLang === 'en' ? ' Rest Space, no event' : ' จุดพักผ่อน ไม่มีเหตุการณ์'; }
                else if (sp.type === 'treasure') { iconHtml = '📦'; accessibleName += currentLang === 'en' ? ' Treasure Chest' : ' หีบสมบัติ'; }
                else if (sp.type === 'forward') { iconHtml = '🚀'; accessibleName += currentLang === 'en' ? ' Warp' : ' วาร์ป'; }
                else if (sp.type === 'trap') { iconHtml = '🕳️'; accessibleName += currentLang === 'en' ? ' Trap' : ' หลุมพราง'; }
                else if (sp.type === 'water') { iconHtml = '🌊'; accessibleName += currentLang === 'en' ? ' Rapids' : ' น้ำเชี่ยว'; }
                else if (sp.type === 'ghost') { iconHtml = '👻'; accessibleName += currentLang === 'en' ? ' Ghost' : ' ผีหลอก'; }
                else if (sp.type === 'warp') { iconHtml = '🌀'; accessibleName += currentLang === 'en' ? ' Cyclone' : ' ไซโคลน'; }
                else if (sp.type === 'key') { iconHtml = '🔑'; accessibleName += currentLang === 'en' ? ' Mystery Box' : ' กล่องลึกลับ'; }
                else if (sp.type === 'door') { iconHtml = '🚪'; accessibleName += currentLang === 'en' ? ` Shortcut Door to space ${sp.dest}` : ` ประตูทางลัด ไปช่อง ${sp.dest}`; }
                else if (sp.type === 'bonus') { iconHtml = '🎁'; accessibleName += currentLang === 'en' ? ' Windfall' : ' ลาภลอย'; }
                else if (sp.type === 'secret') { iconHtml = '🔮'; accessibleName += currentLang === 'en' ? ' Secret Event' : ' เหตุการณ์ลับ'; }
            }

            cell.setAttribute('aria-label', accessibleName);
            cell.innerHTML = `<span class="cell-num">${num}</span> <span class="cell-icon">${iconHtml}</span> <div class="pawns-holder" id="pawns-holder-${num}"></div>`;
            grid.appendChild(cell);
        }
    }
}

// Update Gameplay UI & Render Pawns
function updateGameUI() {
    if (!gameState) return;

    const pOrder = gameState.playerOrder || Object.keys(gameState.players);
    const playersArr = pOrder.map(pId => [pId, gameState.players[pId]]);
    const currentTurnKey = playersArr[gameState.turnIndex][0];
    const currentTurnPlayer = playersArr[gameState.turnIndex][1];

    document.getElementById('game-status-bar').textContent = currentLang === 'en'
        ? `Current Turn: ${currentTurnPlayer.animal.icon} ${currentTurnPlayer.name}`
        : `ถึงเทิร์นของ: ${currentTurnPlayer.animal.icon} ${currentTurnPlayer.name}`;

    const rollBtn = document.getElementById('btn-roll-dice');
    if (currentTurnKey === myPlayerId && !currentTurnPlayer.isBot && !gameState.turnExecuting) {
        rollBtn.disabled = false;
        
        // Auto Play Timer (40 seconds)
        if (!myTurnTimer) {
            myTurnTimer = setTimeout(() => {
                if (!document.getElementById('btn-roll-dice').disabled) {
                    window.handleRollDice(true); // Auto roll and auto use items
                }
            }, 40000);
        }
    } else {
        rollBtn.disabled = true;
        
        // Clear Timer if not my turn
        if (myTurnTimer) {
            clearTimeout(myTurnTimer);
            myTurnTimer = null;
        }
    }

    if (gameState.status === 'playing' && currentTurnPlayer) {
        const turnUniqueId = `${currentRoomId}_${gameState.turnIndex}_${currentTurnKey}`;
        if (lastAnnouncedTurnKey !== turnUniqueId) {
            lastAnnouncedTurnKey = turnUniqueId;
            const turnMsg = currentLang === 'en'
                ? `It is ${currentTurnPlayer.name}'s turn`
                : `ถึงเทิร์นของ ${currentTurnPlayer.name}`;
            announceSR(turnMsg);
            
            // Audio System: Sync Turn & ABC Rule
            playAudio('turn.mp3').then(() => {
                if (currentTurnKey === myPlayerId && !currentTurnPlayer.isBot) {
                    playAudio('abc.mp3');
                }
            });
            
            // Focus the roll button when it is the local player's turn, after announcement
            if (currentTurnKey === myPlayerId && !currentTurnPlayer.isBot && !rollBtn.disabled) {
                const duration = Math.max(1500, turnMsg.length * 50) + 200;
                setTimeout(() => {
                    const btn = document.getElementById('btn-roll-dice');
                    if (btn && !btn.disabled && document.activeElement !== btn) {
                        btn.focus();
                    }
                }, duration);
            }
        }
    }

    const cardsContainer = document.getElementById('player-status-cards');
    cardsContainer.innerHTML = '';
    
    // Group all statuses into a single readable object for Screen Readers
    let allStatusText = currentLang === 'en' ? "All players status: " : "สถานะผู้เล่นทั้งหมด: ";
    playersArr.forEach(([pId, p], index) => {
        if (currentLang === 'en') {
            allStatusText += `${p.name} Space ${p.pos} Keys ${p.keys} Armor ${p.armor}`;
        } else {
            allStatusText += `${p.name} ช่อง ${p.pos} กุญแจ ${p.keys} เกราะ ${p.armor}`;
        }
        if (index < playersArr.length - 1) allStatusText += ", ";
    });
    
    cardsContainer.setAttribute('aria-label', allStatusText);
    cardsContainer.setAttribute('tabindex', '0');
    cardsContainer.setAttribute('role', 'group');

    // Render visual cards for sighted users and hide them from screen reader to prevent duplicate reading
    const spaceLabel = currentLang === 'en' ? 'Space' : 'ช่อง';
    playersArr.forEach(([pId, p]) => {
        const card = document.createElement('div');
        card.className = `p-status-card ${pId === currentTurnKey ? 'active-turn' : ''}`;
        card.setAttribute('aria-hidden', 'true');
        card.innerHTML = `
            <div><strong>${p.animal.icon} ${p.name}</strong></div>
            <div>${spaceLabel} ${p.pos} 🔑 ${p.keys} 🛡️ ${p.armor}</div>
        `;
        cardsContainer.appendChild(card);
    });

    document.querySelectorAll('.pawns-holder').forEach(h => h.innerHTML = '');
    playersArr.forEach(([pId, p]) => {
        const holder = document.getElementById(`pawns-holder-${p.pos}`);
        if (holder) {
            const token = document.createElement('div');
            token.className = 'pawn-token';
            token.innerHTML = `<span class="pawn-name-tag">${p.name}</span>${p.animal.icon}`;
            holder.appendChild(token);
        }
    });

    // --- VISUAL LAYER: HIGHLIGHT CURRENT PLAYER ---
    if (gameState.status === 'playing' && currentTurnPlayer) {
        VisualController.highlightCurrentPlayer(currentTurnKey, currentTurnPlayer, currentTurnPlayer.pos);
    }
    // ----------------------------------------------

    // Host Bot Turn Engine
    if (isHost && currentTurnPlayer.isBot && gameState.status === 'playing' && !gameState.turnExecuting) {
        // Lock executing flag
        update(ref(db), { [`games/Adventure80/rooms/${currentRoomId}/turnExecuting`]: true });
        setTimeout(() => { executeTurnAsync(currentTurnKey, true); }, 1500);
    }
}

// Emits an action and pauses briefly so all clients process audio sequentially
async function syncActionEmit(msg, audioKeys = []) {
    const ts = Date.now() + Math.random();
    await update(ref(db), { [`games/Adventure80/rooms/${currentRoomId}/lastAction`]: { msg, ts, audioKeys } });
    await delayAsync(2200); // Wait long enough for Screen Readers and Web Audio to play the sequence
}

async function syncStateDB(pId, updatesObj) {
    const dbUpdates = {};
    for (const [k, v] of Object.entries(updatesObj)) {
        dbUpdates[`games/Adventure80/rooms/${currentRoomId}/players/${pId}/${k}`] = v;
    }
    await update(ref(db), dbUpdates);
}

// Handle Roll Dice Button Click (Human)
window.handleRollDice = function(isAuto = false) {
    const btn = document.getElementById('btn-roll-dice');
    btn.disabled = true;
    if (myTurnTimer) {
        clearTimeout(myTurnTimer);
        myTurnTimer = null;
    }
    update(ref(db), { [`games/Adventure80/rooms/${currentRoomId}/turnExecuting`]: true });
    executeTurnAsync(myPlayerId, isAuto);
};

// Event Chain Engine (Core Game Logic Refactored)
async function executeTurnAsync(pId, isAuto = false) {
    const pData = gameState.players[pId];
    const autoMode = pData.isBot || isAuto;
    const diceRoll = Math.floor(Math.random() * 6) + 1;
    document.getElementById('dice-visual').textContent = diceRoll;
    
    await syncActionEmit(`${pData.name} ทอยลูกเต๋าได้ ${diceRoll}`, ['dice.mp3']);
    
    let currentPos = pData.pos;
    let targetPos = currentPos + diceRoll;
    if (targetPos > 80) targetPos = 80 - (targetPos - 80); // Bounce back rule

    await syncActionEmit(`${pData.name} เดินจากช่อง ${currentPos} ไปยังช่อง ${targetPos}`, ['walk.mp3']);
    
    currentPos = targetPos;
    await syncStateDB(pId, { pos: currentPos });

    if (currentPos === 80) {
        await handleWinGame(pId);
        return;
    }

    let chainCount = 0;
    let visited = new Set();
    visited.add(currentPos);

    // Event Chain Loop: Max 3 events limit
    while (chainCount < 3) {
        const sp = gameState.boardConfig ? gameState.boardConfig[currentPos] : null;
        if (!sp) {
            // ไม่ใช่ช่องพิเศษ จบเทิร์น
            break;
        }

        let typeNameTH = 'พิเศษ';
        let landingAudio = null;
        if (sp.type === 'rest') { typeNameTH = 'จุดพักผ่อน'; landingAudio = 'sabuy.mp3'; }
        else if (sp.type === 'treasure') { typeNameTH = 'หีบสมบัติ'; landingAudio = 'box1.mp3'; }
        else if (sp.type === 'forward') { typeNameTH = 'วาร์ป'; landingAudio = 'forward.mp3'; }
        else if (sp.type === 'trap') { typeNameTH = 'หลุมพราง'; landingAudio = 'hole.mp3'; }
        else if (sp.type === 'water') { typeNameTH = 'น้ำเชี่ยว'; landingAudio = 'water.mp3'; }
        else if (sp.type === 'ghost') { typeNameTH = 'ผีหลอก'; landingAudio = 'ghost.mp3'; }
        else if (sp.type === 'warp') { typeNameTH = 'ไซโคลน'; landingAudio = 'cyclone.mp3'; }
        else if (sp.type === 'key') { typeNameTH = 'กล่องลึกลับ'; landingAudio = 'box2.mp3'; }
        else if (sp.type === 'door') { typeNameTH = 'ประตูทางลัด'; landingAudio = 'door1.mp3'; }
        else if (sp.type === 'bonus') { typeNameTH = 'ลาภลอย'; landingAudio = 'box4.mp3'; }
        else if (sp.type === 'secret') { typeNameTH = 'เหตุการณ์ลับ'; landingAudio = 'box3.mp3'; }

        await syncActionEmit(`${pData.name}ตกช่อง ${currentPos} เป็นช่อง${typeNameTH}`, landingAudio ? [landingAudio] : []);

        // Handle Effect
        if (sp.type === 'rest') {
            const restNames = ["แคมป์ไฟอบอุ่น", "โอเอซิสสงบเงียบ", "กระท่อมร้างกลางป่า", "ใต้ต้นไม้ใหญ่", "ศาลาพักใจ"];
            const rName = restNames[Math.floor(Math.random() * restNames.length)];
            await syncActionEmit(`${pData.name} มาถึงจุดพักผ่อน ${rName} ช่อง ${currentPos} ไม่มีเหตุการณ์พิเศษ จบการเดินทาง`);
            break; // Rule 5: Ends event chain immediately
        }
        
        let movedToNewSpace = false;

        if (sp.type === 'treasure') {
            if (sp.charges > 0) {
                sp.charges--;
                await update(ref(db), { [`games/Adventure80/rooms/${currentRoomId}/boardConfig/${currentPos}/charges`]: sp.charges });
                pData.armor++;
                await syncStateDB(pId, { armor: pData.armor });
                await syncActionEmit(`พบหีบสมบัติ เปิดหีบ พบเกราะศักดิ์สิทธิ์ ได้รับเกราะศักดิ์สิทธิ์ 1 ชิ้น`, ['akey3.mp3']);
            } else {
                await syncActionEmit(`พบหีบสมบัติ แต่หีบถูกเปิดไปแล้ว เหลือเพียงเศษฝุ่น ไม่ได้อะไรเลย`, ['akey2.mp3']);
            }
            break; // Event chain ends as player did not change location
        } else if (sp.type === 'key') {
            if (sp.charges > 0) {
                sp.charges--;
                await update(ref(db), { [`games/Adventure80/rooms/${currentRoomId}/boardConfig/${currentPos}/charges`]: sp.charges });
                pData.keys++;
                await syncStateDB(pId, { keys: pData.keys });
                await syncActionEmit(`พบกล่องลึกลับ เปิดกล่อง พบกุญแจโบราณ ได้รับกุญแจโบราณ 1 ดอก`, ['akey1.mp3']);
            } else {
                await syncActionEmit(`พบกล่องลึกลับ แต่กล่องถูกเปิดไปแล้ว เหลือเพียงเศษฝุ่น ไม่ได้อะไรเลย`, ['akey2.mp3']);
            }
            break; // Does not change location
        } else if (sp.type === 'door') {
            const doorNames = ["ประตูมิติ", "ประตูกาลเวลา", "ประตูศิลาโบราณ", "ประตูเวทมนตร์", "ประตูลับ"];
            const dName = doorNames[Math.floor(Math.random() * doorNames.length)];
            if (pData.keys > 0) {
                let useKey = true;
                if (!autoMode && pId === myPlayerId) {
                    if (currentLang === 'en') {
                        const doorEnMap = {
                            'ประตูมิติ': 'Dimensional Door',
                            'ประตูกาลเวลา': 'Time Door',
                            'ประตูศิลาโบราณ': 'Ancient Stone Door',
                            'ประตูเวทมนตร์': 'Magic Door',
                            'ประตูลับ': 'Secret Door'
                        };
                        const dNameEn = doorEnMap[dName] || 'Shortcut Door';
                        useKey = await showModalAsync(`Found ${dNameEn}!`, 'You have a key. Use it to open the door and warp ahead?', 'Use Key', 'Cancel');
                    } else {
                        useKey = await showModalAsync(`พบ${dName}!`, 'คุณมีกุญแจ ต้องการใช้กุญแจเปิดประตูไปข้างหน้าหรือไม่?', 'ใช้กุญแจ', 'ไม่ใช้');
                    }
                }
                if (useKey) {
                    pData.keys--;
                    await syncStateDB(pId, { keys: pData.keys });
                    await syncActionEmit(`ใช้กุญแจเปิด${dName}สำเร็จ วาร์ปไปยังช่อง ${sp.dest}`, ['door2.mp3', 'walk.mp3']);
                    currentPos = sp.dest;
                    movedToNewSpace = true;
                } else {
                    await syncActionEmit(`เลือกที่จะไม่ใช้กุญแจ ประตูยังคงปิดอยู่`, ['shieldno.mp3']);
                    break;
                }
            } else {
                await syncActionEmit(`พบ${dName} แต่ไม่มีกุญแจ ไม่สามารถเปิดได้`);
                break;
            }
        } else if (['trap', 'water', 'ghost', 'warp'].includes(sp.type)) {
            let randVal;
            if (sp.type === 'warp' || sp.type === 'water') {
                randVal = Math.random() > 0.5 ? (Math.floor(Math.random() * 10) + 3) : -(Math.floor(Math.random() * 5) + 2);
            } else {
                randVal = -(Math.floor(Math.random() * 5) + 2);
            }
            let isBad = (randVal < 0);
            if (isBad && pData.armor > 0) {
                let useArmor = true;
                if (!autoMode && pId === myPlayerId) {
                    if (currentLang === 'en') {
                        useArmor = await showModalAsync('Danger Ahead!', `Space ${currentPos} has a bad event. Use Holy Armor to block it?`, 'Use Armor', 'Cancel');
                    } else {
                        useArmor = await showModalAsync('พบอันตราย!', `ช่อง ${currentPos} มีเหตุการณ์ร้าย ใช้เกราะศักดิ์สิทธิ์ป้องกันหรือไม่?`, 'ใช้เกราะป้องกัน', 'ไม่ใช้');
                    }
                }
                if (useArmor) {
                    pData.armor--;
                    await syncStateDB(pId, { armor: pData.armor });
                    await syncActionEmit(`ใช้เกราะศักดิ์สิทธิ์ป้องกันผลเสียสำเร็จ ปลอดภัยแล้ว!`, ['shield.mp3']);
                    break;
                } else {
                    await syncActionEmit(`เลือกไม่ใช้เกราะศักดิ์สิทธิ์ รับผลร้าย ถอยหลัง ${Math.abs(randVal)} ช่อง`, ['shieldno.mp3', 'walk.mp3']);
                    currentPos = Math.max(1, currentPos + randVal);
                    movedToNewSpace = true;
                }
            } else if (isBad) {
                await syncActionEmit(`เกิดเหตุการณ์ร้าย ถอยหลัง ${Math.abs(randVal)} ช่อง`, ['walk.mp3']);
                currentPos = Math.max(1, currentPos + randVal);
                movedToNewSpace = true;
            } else {
                await syncActionEmit(`โชคดี เดินหน้าเพิ่ม ${randVal} ช่อง`, ['forward.mp3', 'walk.mp3']);
                let targetPos = currentPos + randVal;
                if (targetPos > 80) targetPos = 80 - (targetPos - 80);
                currentPos = targetPos;
                movedToNewSpace = true;
            }
        } else if (sp.type === 'forward') {
            let randVal = Math.floor(Math.random() * 10) + 3;
            await syncActionEmit(`พลังพิเศษ เดินหน้า ${randVal} ช่อง`, ['forward.mp3', 'walk.mp3']);
            let targetPos = currentPos + randVal;
            if (targetPos > 80) targetPos = 80 - (targetPos - 80);
            currentPos = targetPos;
            movedToNewSpace = true;
        } else if (sp.type === 'bonus') {
            const randEffect = Math.floor(Math.random() * 3);
            if (randEffect === 0) {
                pData.keys++;
                await syncStateDB(pId, { keys: pData.keys });
                await syncActionEmit(`ลาภลอย! ได้รับกุญแจ 1 ชิ้น`, ['akey1.mp3']);
                break;
            } else if (randEffect === 1) {
                pData.armor++;
                await syncStateDB(pId, { armor: pData.armor });
                await syncActionEmit(`ลาภลอย! ได้รับเกราะศักดิ์สิทธิ์ 1 ชิ้น`, ['akey3.mp3']);
                break;
            } else {
                let randVal = Math.floor(Math.random() * 10) + 3;
                let actionDesc = `เดินหน้า ${randVal} ช่อง`;
                await syncActionEmit(`ลาภลอย! เกิดการวาร์ป ${actionDesc}`, ['forward.mp3', 'walk.mp3']);
                let targetPos = currentPos + randVal;
                if (targetPos > 80) targetPos = 80 - (targetPos - 80);
                currentPos = targetPos;
                movedToNewSpace = true;
            }
        } else if (sp.type === 'secret') {
            const secretNames = ["วิญญาณศักดิ์สิทธิ์", "สัตว์เวทย์นำทาง", "ภูตแห่งแสง"];
            const sName = secretNames[Math.floor(Math.random() * secretNames.length)];
            let randVal = Math.floor(Math.random() * 10) + 3;
            let actionDesc = `พาเดินหน้า ${randVal} ช่อง`;
            await syncActionEmit(`พบเหตุการณ์ลับ: ${sName} ${actionDesc}`, ['forward.mp3', 'walk.mp3']);
            let targetPos = currentPos + randVal;
            if (targetPos > 80) targetPos = 80 - (targetPos - 80);
            currentPos = targetPos;
            movedToNewSpace = true;
        }

        // Rule 8 & 7: Check newly moved destination, stop if loop, limit 3 chain
        if (movedToNewSpace) {
            await syncStateDB(pId, { pos: currentPos });
            
            if (currentPos === 80) {
                await handleWinGame(pId);
                return;
            }

            if (currentPos === 1) {
                await syncActionEmit(`กลับมายังจุดเริ่มต้น`, ['sabuy.mp3']);
            }

            if (visited.has(currentPos)) {
                await syncActionEmit(`ตรวจสอบช่องใหม่: กลับมาช่อง ${currentPos} ที่เคยผ่านมาแล้ว หยุดการเดินทางเพื่อป้องกันการวนลูป`);
                break;
            }
            
            visited.add(currentPos);
            chainCount++;
            
            if (chainCount >= 3) {
                await syncActionEmit(`ตรวจสอบช่องใหม่: ครบขีดจำกัดเหตุการณ์ 3 ครั้งแล้ว จบการเคลื่อนที่สำหรับเทิร์นนี้ทันที`);
                break; // Stop Chain
            } else {
                await syncActionEmit(`ตรวจสอบช่องใหม่: ไปยังช่อง ${currentPos}`);
                // Loop continues to re-evaluate the new currentPos
            }
        } else {
            // Did not move to a new space, end chain loop
            break;
        }
    }

    // Finished turn
    const pKeys = gameState.playerOrder || Object.keys(gameState.players);
    const nextTurn = (gameState.turnIndex + 1) % pKeys.length;
    await update(ref(db), { 
        [`games/Adventure80/rooms/${currentRoomId}/turnIndex`]: nextTurn,
        [`games/Adventure80/rooms/${currentRoomId}/turnExecuting`]: false
    });
}

// End Game Process
async function handleWinGame(pId) {
    await update(ref(db), { 
        [`games/Adventure80/rooms/${currentRoomId}/status`]: 'ended',
        [`games/Adventure80/rooms/${currentRoomId}/winnerId`]: pId,
        [`games/Adventure80/rooms/${currentRoomId}/turnExecuting`]: false
    });
}

// Accessible Modal Promise wrapper
function showModalAsync(title, desc, confirmText, cancelText) {
    return new Promise((resolve) => {
        const overlay = document.getElementById('modal-overlay');
        document.getElementById('modal-title').textContent = title;
        document.getElementById('modal-desc').textContent = desc;

        const btn1 = document.getElementById('btn-modal-action-1');
        const btn2 = document.getElementById('btn-modal-action-2');

        // จัดการลบแอตทริบิวต์เพื่อป้องกันปัญหา VoiceOver อ่านข้อความซ้ำซ้อนบน iOS
        overlay.removeAttribute('aria-labelledby');
        overlay.removeAttribute('aria-describedby');
        btn1.removeAttribute('aria-labelledby');
        btn1.removeAttribute('aria-describedby');
        btn2.removeAttribute('aria-labelledby');
        btn2.removeAttribute('aria-describedby');

        btn1.textContent = confirmText;
        btn2.textContent = cancelText;

        overlay.style.display = 'flex';
        document.getElementById('modal-title').focus(); // Accessible Focus management

        btn1.onclick = () => { overlay.style.display = 'none'; resolve(true); };
        btn2.onclick = () => { overlay.style.display = 'none'; resolve(false); };
    });
}

// Show Result Screen
function showResultScreen() {
    switchScreen('screen-result', 'result-title');
    
    // ล้างข้อความและสถานะ Live Region เพื่อไม่ให้ค้างไปหน้าจบเกม
    speechQueue = [];
    const srPolite = document.getElementById('sr-polite');
    const srAssertive = document.getElementById('sr-assertive');
    if (srPolite) srPolite.textContent = '';
    if (srAssertive) srAssertive.textContent = '';
    
    const statusBar = document.getElementById('game-status-bar');
    if (statusBar) statusBar.textContent = '';

    const winnerId = gameState.winnerId;
    const winner = gameState.players[winnerId];

    if (currentLang === 'en') {
        document.getElementById('result-winner-text').textContent = `🎉 ${winner.animal.icon} ${winner.name} reached Space 80 first!`;
        announceSR(`The adventure is complete! The winner is ${winner.name}!`, 'assertive');
    } else {
        document.getElementById('result-winner-text').textContent = `🎉 ${winner.animal.icon} ${winner.name} เข้าสู่ช่อง 80 เป็นคนแรก!`;
        announceSR(`การผจญภัยสิ้นสุดลง ผู้ชนะคือ ${winner.name} !`, 'assertive');
    }

    const rankingsBox = document.getElementById('result-rankings');
    rankingsBox.innerHTML = currentLang === 'en' ? '<h3>Final Rankings:</h3>' : '<h3>อันดับการเดินทาง:</h3>';

    const sorted = Object.values(gameState.players).sort((a, b) => b.pos - a.pos);
    const spaceLabel = currentLang === 'en' ? 'Space' : 'ช่อง';
    sorted.forEach((p) => {
        rankingsBox.innerHTML += `<p>${p.name} ${spaceLabel} ${p.pos}</p>`;
    });
    
    // --- VISUAL LAYER: RESULT EFFECTS ---
    setTimeout(() => {
        VisualController.playResultEffects(document.getElementById('screen-result'));
    }, 100);
    // ------------------------------------
}

// Leave Room / Return to Main Menu
window.leaveRoom = function() {
    if (currentRoomId && myPlayerId) {
        if (isHost) {
            remove(ref(db, `games/Adventure80/rooms/${currentRoomId}`));
        } else {
            update(ref(db, `games/Adventure80/rooms/${currentRoomId}/players/${myPlayerId}`), {
                isBot: true,
                name: "บอท" + myPlayerName
            });
        }
    }
    stopBGM();
    myPlayerId = null;
    currentRoomId = null;
    isHost = false;
    gameState = null;
    lastAnnouncedTurnKey = null;
    isShowingWinnerScene = false;
    previousHumanCount = null;
    previousBotCount = null;
    isStartingGame = false;
    if (myTurnTimer) {
        clearTimeout(myTurnTimer);
        myTurnTimer = null;
    }
    document.getElementById('lobby-room-section').style.display = 'none';
    document.getElementById('lobby-menu-section').style.display = 'block';
};

window.returnToLobbyOrMain = function() {
    window.leaveRoom();
    switchScreen('screen-lobby', 'lobby-heading');
};

// ==================================================
// PURE VISUAL CONTROLLER LAYER
// โดยไม่แตะต้อง Accessibility (ARIA) และ Game Logic เดิม
// ==================================================
const VisualController = {
    isReducedMotion: () => window.matchMedia('(prefers-reduced-motion: reduce)').matches,

    init: function() {
        if (document.getElementById('vc-styles')) return;
        const style = document.createElement('style');
        style.id = 'vc-styles';
        style.innerHTML = `
            .visual-effect {
                animation-fill-mode: forwards;
                pointer-events: none;
            }
            .effect-walk { animation: vc-fadeUp 1s ease-out; }
            .effect-forward { color: #0f0; animation: vc-slideRight 1.5s ease-out; }
            .effect-backward { color: #f00; animation: vc-slideLeft 1.5s ease-out; }
            .effect-water { animation: vc-wave 2s ease-in-out; }
            .effect-cyclone { animation: vc-spin 2s linear; }
            .effect-trap { animation: vc-drop 2s ease-in; }
            .effect-ghost { animation: vc-float 2s ease-in-out; text-shadow: 0 0 10px purple; }
            .effect-warp { animation: vc-pulse 2s ease-in-out; }
            .effect-secret { animation: vc-sparkle 2s ease-in-out; }
            .effect-bonus { animation: vc-scaleUp 2s ease-out; }
            .effect-dust { opacity: 0.5; animation: vc-fadeRight 1.5s ease-out; }
            .effect-armor-fly { animation: vc-flyUp 2s ease-out; text-shadow: 0 0 15px gold; }
            .effect-key-fly { animation: vc-flyUp 2s ease-out; text-shadow: 0 0 15px gold; }
            .effect-door-open { animation: vc-flash 2s ease-in-out; }
            .effect-door-closed { animation: vc-shake 1.5s ease-in-out; }
            .effect-rest { animation: vc-glow 2.5s ease-in-out; }
            .effect-holy-shield { animation: vc-shield 2s ease-out; font-size: 3rem; }
            
            .visual-dice-anim { animation: vc-diceRoll 0.8s ease-in-out; display: inline-block; }
            .visual-active-pawn { 
                box-shadow: 0 0 15px 5px rgba(255, 215, 0, 0.8) !important; 
                border-radius: 50%; 
                transform: scale(1.1); 
                transition: all 0.3s; 
            }
            .visual-bounce { animation: vc-bounce 0.5s ease-out; }
            .visual-scale { animation: vc-scaleUp 1s ease-in-out; }
            .visual-rotate { animation: vc-spin 1s linear; }
            .visual-drop { animation: vc-drop 1s ease-in; }
            .visual-shake { animation: vc-shake 1s ease-in-out; }
            
            .visual-confetti {
                position: absolute; top: -10px; width: 10px; height: 10px; background: gold;
                animation: vc-fall 3s linear infinite; pointer-events: none; z-index: 2000;
            }
            .effect-trophy-huge { font-size: 6rem; position: absolute; left: 50%; top: 30%; transform: translate(-50%, -50%); animation: vc-scaleUpBounce 2s ease-out; }
            .effect-sparkle { font-size: 4rem; position: absolute; left: 50%; top: 30%; transform: translate(-50%, -50%); animation: vc-spin 4s linear infinite; }
            .effect-fireworks { font-size: 5rem; position: absolute; left: 50%; top: 20%; transform: translateX(-50%); animation: vc-flash 2s infinite; }
            .visual-stagger-in { opacity: 0; animation: vc-fadeUp 0.5s forwards; }
            
            @keyframes vc-fadeUp { 0% { opacity: 0; transform: translateY(20px); } 100% { opacity: 1; transform: translateY(0); } }
            @keyframes vc-slideRight { 0% { opacity: 0; transform: translateX(-20px); } 100% { opacity: 1; transform: translateX(20px); } }
            @keyframes vc-slideLeft { 0% { opacity: 0; transform: translateX(20px); } 100% { opacity: 1; transform: translateX(-20px); } }
            @keyframes vc-fadeRight { 0% { opacity: 0; transform: translateX(0); } 100% { opacity: 0; transform: translateX(20px); } }
            @keyframes vc-wave { 0%, 100% { transform: translateY(0); } 50% { transform: translateY(-15px) scale(1.2); } }
            @keyframes vc-spin { 100% { transform: rotate(360deg); } }
            @keyframes vc-drop { 0% { transform: scale(1); opacity: 1; } 100% { transform: scale(0.1) translateY(50px); opacity: 0; } }
            @keyframes vc-float { 0%, 100% { transform: translateY(0); } 50% { transform: translateY(-20px); } }
            @keyframes vc-pulse { 0%, 100% { transform: scale(1); opacity: 1; } 50% { transform: scale(1.5); opacity: 0.7; } }
            @keyframes vc-sparkle { 0%, 100% { opacity: 1; transform: scale(1); } 50% { opacity: 0.5; transform: scale(1.3); } }
            @keyframes vc-scaleUp { 0% { transform: scale(0.5); opacity: 0; } 100% { transform: scale(1.2); opacity: 1; } }
            @keyframes vc-flyUp { 0% { transform: translateY(20px); opacity: 0; } 100% { transform: translateY(-30px); opacity: 1; } }
            @keyframes vc-flash { 0%, 100% { opacity: 1; } 50% { opacity: 0.2; } }
            @keyframes vc-shake { 0%, 100% { transform: translateX(0); } 25% { transform: translateX(-5px); } 75% { transform: translateX(5px); } }
            @keyframes vc-glow { 0%, 100% { text-shadow: 0 0 5px orange; } 50% { text-shadow: 0 0 20px red; } }
            @keyframes vc-shield { 0% { transform: scale(0); opacity: 0; } 50% { transform: scale(1.5); opacity: 1; } 100% { transform: scale(1); opacity: 0; } }
            @keyframes vc-diceRoll { 0% { transform: rotate(0) scale(1); } 50% { transform: rotate(180deg) scale(1.2); } 100% { transform: rotate(360deg) scale(1); } }
            @keyframes vc-bounce { 0%, 100% { transform: translateY(0); } 50% { transform: translateY(-15px); } }
            @keyframes vc-scaleUpBounce { 0% { transform: translate(-50%, -50%) scale(0); } 80% { transform: translate(-50%, -50%) scale(1.2); } 100% { transform: translate(-50%, -50%) scale(1); } }
            @keyframes vc-fall { 0% { transform: translateY(0) rotate(0); opacity: 1; } 100% { transform: translateY(200px) rotate(360deg); opacity: 0; } }
        `;
        document.head.appendChild(style);
    },

    createEffect: function(container, content, className, duration) {
        if (this.isReducedMotion()) return;
        const wrapper = document.createElement('div');
        wrapper.setAttribute('aria-hidden', 'true'); // ACCESSIBILITY: New DOM purely visual
        wrapper.style.position = 'absolute';
        wrapper.style.left = '50%';
        wrapper.style.top = '50%';
        wrapper.style.pointerEvents = 'none';
        wrapper.style.zIndex = '1000';
        
        const el = document.createElement('div');
        el.className = 'visual-effect ' + className;
        el.innerHTML = content;
        wrapper.style.marginLeft = '-1rem';
        wrapper.style.marginTop = '-1rem';
        wrapper.style.width = '2rem';
        wrapper.style.height = '2rem';
        wrapper.style.display = 'flex';
        wrapper.style.alignItems = 'center';
        wrapper.style.justifyContent = 'center';
        wrapper.style.fontSize = '2rem';

        wrapper.appendChild(el);

        if (container) container.appendChild(wrapper);
        if (duration) setTimeout(() => { if (wrapper.parentNode) wrapper.parentNode.removeChild(wrapper); }, duration);
        return wrapper;
    },

    animateDice: function() {
        if (this.isReducedMotion()) return;
        const diceEl = document.getElementById('dice-visual');
        if (diceEl) {
            diceEl.classList.remove('visual-dice-anim');
            void diceEl.offsetWidth; // trigger reflow
            diceEl.classList.add('visual-dice-anim');
            setTimeout(() => diceEl.classList.remove('visual-dice-anim'), 800);
        }
    },

    highlightCurrentPlayer: function(pId, pData, currentPos) {
        document.querySelectorAll('.pawn-token').forEach(p => p.classList.remove('visual-active-pawn'));
        const holder = document.getElementById(`pawns-holder-${currentPos}`);
        if (holder) {
            const pawns = holder.querySelectorAll('.pawn-token');
            pawns.forEach(pawn => {
                if (pawn.textContent.includes(pData.name)) {
                    pawn.classList.add('visual-active-pawn');
                }
            });
        }
    },

    triggerFromMessage: function(msg, pData) {
        if (this.isReducedMotion()) return;
        
        let targetPos = pData ? pData.pos : 1;
        const match = msg.match(/ช่อง\s*(\d+)/);
        if (match) targetPos = parseInt(match[1]);

        const cell = document.getElementById(`cell-${targetPos}`);
        if (!cell) return;
        
        let activePawn = null;
        const pawnHolder = document.getElementById(`pawns-holder-${targetPos}`);
        if (pawnHolder && pData) {
            Array.from(pawnHolder.querySelectorAll('.pawn-token')).forEach(p => {
                if (p.textContent.includes(pData.name)) activePawn = p;
            });
        }

        // Event Matching Visuals
        if (msg.includes('ทอยลูกเต๋าได้')) {
            this.animateDice();
        } else if (msg.includes('เดินจากช่อง') || msg.includes('ถอยหลัง') || msg.includes('เดินหน้า') || msg.includes('กลับมายังจุดเริ่มต้น')) {
            if (msg.includes('ถอยหลัง')) {
                this.createEffect(cell, '⬅️', 'effect-backward', 1500);
            } else if (msg.includes('เดินหน้า') || msg.includes('พลังพิเศษ') || msg.includes('พาเดินหน้า')) {
                this.createEffect(cell, '➡️', 'effect-forward', 1500);
            } else {
                this.createEffect(cell, '👣', 'effect-walk', 1000);
            }
            if (activePawn) {
                activePawn.classList.add('visual-bounce');
                setTimeout(() => activePawn.classList.remove('visual-bounce'), 500);
            }
        } else if (msg.includes('น้ำเชี่ยว')) {
            this.createEffect(cell, '🌊', 'effect-water', 2000);
            if (activePawn) { activePawn.classList.add('visual-scale'); setTimeout(() => activePawn.classList.remove('visual-scale'), 1000); }
        } else if (msg.includes('ไซโคลน')) {
            this.createEffect(cell, '🌪️', 'effect-cyclone', 2000);
            if (activePawn) { activePawn.classList.add('visual-rotate'); setTimeout(() => activePawn.classList.remove('visual-rotate'), 1000); }
        } else if (msg.includes('หลุมพราง') || msg.includes('เหตุการณ์ร้าย')) {
            this.createEffect(cell, '🕳️', 'effect-trap', 2000);
            if (activePawn) { activePawn.classList.add('visual-drop'); setTimeout(() => activePawn.classList.remove('visual-drop'), 1000); }
        } else if (msg.includes('ผีหลอก')) {
            this.createEffect(cell, '👻', 'effect-ghost', 2000);
            if (activePawn) { activePawn.classList.add('visual-shake'); setTimeout(() => activePawn.classList.remove('visual-shake'), 1000); }
        } else if (msg.includes('วาร์ปไปยังช่อง')) {
            this.createEffect(cell, '🌀', 'effect-warp', 2000);
        } else if (msg.includes('เหตุการณ์ลับ')) {
            this.createEffect(cell, '✨', 'effect-secret', 2000);
        } else if (msg.includes('ลาภลอย')) {
            this.createEffect(cell, '✨✨✨', 'effect-bonus', 2000);
        } else if (msg.includes('หีบสมบัติ')) {
            if (msg.includes('เหลือเพียงเศษฝุ่น')) this.createEffect(cell, '💨', 'effect-dust', 1500);
            else this.createEffect(cell, '🛡️✨', 'effect-armor-fly', 2000);
        } else if (msg.includes('กล่องลึกลับ')) {
             if (msg.includes('เหลือเพียงเศษฝุ่น')) this.createEffect(cell, '💨', 'effect-dust', 1500);
             else this.createEffect(cell, '🔑✨', 'effect-key-fly', 2000);
        } else if (msg.includes('ประตู')) {
            if (msg.includes('ใช้กุญแจเปิด')) this.createEffect(cell, '🔑✨🚪', 'effect-door-open', 2000);
            else if (msg.includes('ไม่มีกุญแจ') || msg.includes('ประตูยังคงปิด')) this.createEffect(cell, '🚪❌', 'effect-door-closed', 1500);
        } else if (msg.includes('จุดพักผ่อน')) {
            this.createEffect(cell, '🔥', 'effect-rest', 2500);
        } else if (msg.includes('ใช้เกราะศักดิ์สิทธิ์ป้องกันผลเสียสำเร็จ')) {
            this.createEffect(cell, '🛡️✨', 'effect-holy-shield', 2000);
        }
    },

    playWinnerEffects: function(container) {
        if (this.isReducedMotion()) return;
        this.createEffect(container, '🏆', 'effect-trophy-huge', 5000);
        this.createEffect(container, '✨', 'effect-sparkle', 5000);
        for(let i=0; i<30; i++) {
            const conf = document.createElement('div');
            conf.className = 'visual-confetti';
            conf.setAttribute('aria-hidden', 'true');
            conf.style.left = (Math.random() * 100) + '%';
            conf.style.animationDelay = (Math.random() * 2) + 's';
            conf.style.background = ['gold', 'red', 'blue', 'green', 'purple'][Math.floor(Math.random()*5)];
            container.appendChild(conf);
            setTimeout(() => { if (conf.parentNode) conf.parentNode.removeChild(conf); }, 5000);
        }
    },
    
    playResultEffects: function(container) {
        if (this.isReducedMotion()) return;
        this.createEffect(container, '🎇', 'effect-fireworks', 4000);
        const cards = container.querySelectorAll('p');
        cards.forEach((c, idx) => {
            c.classList.add('visual-stagger-in');
            c.style.animationDelay = (idx * 0.2) + 's';
        });
    },
    
    playStartEffects: function(container) {
        if (this.isReducedMotion()) return;
        this.createEffect(container, '🏕️', 'effect-trophy-huge', 4000);
        this.createEffect(container, '✨', 'effect-sparkle', 4000);
    }
};

// Initialize Application on DOM Ready
document.addEventListener('DOMContentLoaded', () => {
    applyLanguageUI();
    initRoomListListener();
    VisualController.init(); // Initialize Visual Layer

    const nameInput = document.getElementById('player-name-input');
    const confirmBtn = document.getElementById('btn-confirm-name');

    if (nameInput && confirmBtn) {
        nameInput.addEventListener('input', () => {
            if (nameInput.value.trim().length > 0) {
                confirmBtn.disabled = false;
            } else {
                confirmBtn.disabled = true;
            }
        });

        const handleEnterConfirm = (e) => {
            const isEnterKey = e.key === 'Enter' || e.code === 'Enter' || e.code === 'NumpadEnter' || e.keyCode === 13;
            if (isEnterKey && !e.isComposing) {
                e.preventDefault();
                e.stopPropagation();
                if (nameInput.value.trim().length > 0 && !confirmBtn.disabled) {
                    window.confirmNameAndEnterLobby();
                }
            }
        };

        nameInput.addEventListener('keydown', handleEnterConfirm);

        // Support form submission if input is wrapped in a form on mobile Safari
        if (nameInput.form) {
            nameInput.form.addEventListener('submit', (e) => {
                e.preventDefault();
                if (nameInput.value.trim().length > 0 && !confirmBtn.disabled) {
                    window.confirmNameAndEnterLobby();
                }
            });
        }
    }

    // Keyboard Shortcuts
    document.addEventListener('keydown', (e) => {
        // Alt + R: Roll Dice Shortcut
        if (e.altKey && e.key.toLowerCase() === 'r') {
            const rollBtn = document.getElementById('btn-roll-dice');
            if (rollBtn && !rollBtn.disabled) {
                e.preventDefault();
                if (!e.repeat) {
                    window.handleRollDice(false);
                }
            }
        }
        
        // Alt + A: Announce All Players' Status
        if (e.altKey && e.key.toLowerCase() === 'a') {
            e.preventDefault();
            if (gameState && gameState.players) {
                let allStatusText = currentLang === 'en' ? "All players status: " : "สถานะผู้เล่นทั้งหมด: ";
                const pOrder = gameState.playerOrder || Object.keys(gameState.players);
                pOrder.forEach((pId, index) => {
                    const p = gameState.players[pId];
                    if (currentLang === 'en') {
                        allStatusText += `${p.name} Space ${p.pos} Keys ${p.keys} Armor ${p.armor}`;
                    } else {
                        allStatusText += `${p.name} ช่อง ${p.pos} กุญแจ ${p.keys} เกราะ ${p.armor}`;
                    }
                    if (index < pOrder.length - 1) allStatusText += ", ";
                });
                announceSR(allStatusText, 'polite');
            }
        }
    });
});