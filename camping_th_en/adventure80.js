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
// BILINGUAL DICTIONARY (TH / EN) & SYNC HELPER
// ==================================================
const LANG_STORAGE_KEY = 'adv80_lang';

function getCurrentLang() {
    if (window.ADV80_LANG === 'th' || window.ADV80_LANG === 'en') {
        return window.ADV80_LANG;
    }
    try {
        const saved = localStorage.getItem(LANG_STORAGE_KEY);
        if (saved === 'th' || saved === 'en') return saved;
    } catch (e) {}
    const sysLang = (navigator.language || navigator.userLanguage || 'en').toLowerCase();
    return sysLang.startsWith('th') ? 'th' : 'en';
}

const I18N_JS = {
    th: {
        defaultPlayerName: 'ผู้เล่นใหม่',
        defaultHostName: 'ผู้เล่น 1',
        defaultJoinName: 'ผู้เล่นเข้าใหม่',
        botPrefix: 'บอท',
        botTag: '(บอท)',
        botNames: ['บอทจอมทัพ', 'บอทน้ำพุ', 'บอทสายรุ้ง', 'บอทวายุ', 'บอทปัญญา'],
        welcomeLobbyAnnounce: (name) => `ยินดีต้อนรับ ${name} เข้าสู่ล็อบบี้การผจญภัย`,
        manualOpened: 'เปิดคู่มือการเล่นแล้ว',
        manualClosed: 'ปิดคู่มือการเล่นแล้ว',
        roomItemLabel: (rId, pCount) => `<span><strong>ห้อง ${rId}</strong> (${pCount}/6 คน)</span>`,
        noRoomsWaiting: 'ไม่มีห้องที่กำลังรอผู้เล่นอยู่...',
        roomCreatedAnnounce: (rId) => `สร้างห้องสำเร็จ รหัสห้องคือ ${rId}`,
        roomJoinedAnnounce: (rId) => `เข้าร่วมห้อง ${rId} เรียบร้อยแล้ว`,
        hostLeftAnnounce: 'เจ้าของห้องออกจากเกม ห้องถูกปิด',
        playerDisconnectedBot: (name) => `${name} หลุดจากห้อง แปลงเป็นบอท`,
        startGameAnnounce: 'การผจญภัยเริ่มต้นขึ้นแล้ว เตรียมตัวให้พร้อม!',
        lobbyRoomTitle: (rId) => `ห้องเกม: ${rId}`,
        cellPrefix: (num) => `ช่อง ${num}`,
        cellStart: ' จุดเริ่มต้น',
        cellFinish: ' เส้นชัยปราสาท',
        cellRest: ' จุดพักผ่อน ไม่มีเหตุการณ์',
        cellTreasure: ' หีบสมบัติ',
        cellForward: ' วาร์ป',
        cellTrap: ' หลุมพราง',
        cellWater: ' น้ำเชี่ยว',
        cellGhost: ' ผีหลอก',
        cellWarp: ' ไซโคลน',
        cellKey: ' กล่องลึกลับ',
        cellDoor: (dest) => ` ประตูทางลัด ไปช่อง ${dest}`,
        cellBonus: ' ลาภลอย',
        cellSecret: ' เหตุการณ์ลับ',
        turnStatusBar: (icon, name) => `ถึงเทิร์นของ: ${icon} ${name}`,
        turnAnnounce: (name) => `ถึงเทิร์นของ ${name}`,
        allPlayersStatusHeader: 'สถานะผู้เล่นทั้งหมด: ',
        playerStatusItem: (p) => `${p.name} ช่อง ${p.pos} กุญแจ ${p.keys} เกราะ ${p.armor}`,
        playerCardStats: (p) => `ช่อง ${p.pos} 🔑 ${p.keys} 🛡️ ${p.armor}`,
        spaceTypes: {
            special: 'พิเศษ',
            rest: 'จุดพักผ่อน',
            treasure: 'หีบสมบัติ',
            forward: 'วาร์ป',
            trap: 'หลุมพราง',
            water: 'น้ำเชี่ยว',
            ghost: 'ผีหลอก',
            warp: 'ไซโคลน',
            key: 'กล่องลึกลับ',
            door: 'ประตูทางลัด',
            bonus: 'ลาภลอย',
            secret: 'เหตุการณ์ลับ'
        },
        restNames: ['แคมป์ไฟอบอุ่น', 'โอเอซิสสงบเงียบ', 'กระท่อมร้างกลางป่า', 'ใต้ต้นไม้ใหญ่', 'ศาลาพักใจ'],
        doorNames: ['ประตูมิติ', 'ประตูกาลเวลา', 'ประตูศิลาโบราณ', 'ประตูเวทมนตร์', 'ประตูลับ'],
        secretNames: ['วิญญาณศักดิ์สิทธิ์', 'สัตว์เวทย์นำทาง', 'ภูตแห่งแสง'],
        modalDoorTitle: (dName) => `พบ${dName}!`,
        modalDoorDesc: 'คุณมีกุญแจ ต้องการใช้กุญแจเปิดประตูไปข้างหน้าหรือไม่?',
        modalDoorConfirm: 'ใช้กุญแจ',
        modalDoorCancel: 'ไม่ใช้',
        modalDangerTitle: 'พบอันตราย!',
        modalDangerDesc: (pos) => `ช่อง ${pos} มีเหตุการณ์ร้าย ใช้เกราะศักดิ์สิทธิ์ป้องกันหรือไม่?`,
        modalDangerConfirm: 'ใช้เกราะป้องกัน',
        modalDangerCancel: 'ไม่ใช้',
        resultWinnerText: (icon, name) => `🎉 ${icon} ${name} เข้าสู่ช่อง 80 เป็นคนแรก!`,
        resultWinnerAnnounce: (name) => `การผจญภัยสิ้นสุดลง ผู้ชนะคือ ${name} !`,
        resultRankingsHeader: '<h3>อันดับการเดินทาง:</h3>',
        resultRankingRow: (name, pos) => `<p>${name} ช่อง ${pos}</p>`,
        events: {
            roll_dice: (p) => `${p.name} ทอยลูกเต๋าได้ ${p.dice}`,
            walk_space: (p) => `${p.name} เดินจากช่อง ${p.from} ไปยังช่อง ${p.to}`,
            land_space: (p) => `${p.name}ตกช่อง ${p.pos} เป็นช่อง${I18N_JS.th.spaceTypes[p.spaceType] || p.spaceType}`,
            rest_arrive: (p) => `${p.name} มาถึงจุดพักผ่อน ${I18N_JS.th.restNames[p.restIdx] || ''} ช่อง ${p.pos} ไม่มีเหตุการณ์พิเศษ จบการเดินทาง`,
            treasure_get: () => 'พบหีบสมบัติ เปิดหีบ พบเกราะศักดิ์สิทธิ์ ได้รับเกราะศักดิ์สิทธิ์ 1 ชิ้น',
            treasure_empty: () => 'พบหีบสมบัติ แต่หีบถูกเปิดไปแล้ว เหลือเพียงเศษฝุ่น ไม่ได้อะไรเลย',
            key_get: () => 'พบกล่องลึกลับ เปิดกล่อง พบกุญแจโบราณ ได้รับกุญแจโบราณ 1 ดอก',
            key_empty: () => 'พบกล่องลึกลับ แต่กล่องถูกเปิดไปแล้ว เหลือเพียงเศษฝุ่น ไม่ได้อะไรเลย',
            door_open: (p) => `ใช้กุญแจเปิด${I18N_JS.th.doorNames[p.doorIdx] || 'ประตู'}สำเร็จ วาร์ปไปยังช่อง ${p.dest}`,
            door_skip: () => 'เลือกที่จะไม่ใช้กุญแจ ประตูยังคงปิดอยู่',
            door_nokey: (p) => `พบ${I18N_JS.th.doorNames[p.doorIdx] || 'ประตู'} แต่ไม่มีกุญแจ ไม่สามารถเปิดได้`,
            armor_block: () => 'ใช้เกราะศักดิ์สิทธิ์ป้องกันผลเสียสำเร็จ ปลอดภัยแล้ว!',
            armor_skip_back: (p) => `เลือกไม่ใช้เกราะศักดิ์สิทธิ์ รับผลร้าย ถอยหลัง ${p.steps} ช่อง`,
            bad_back: (p) => `เกิดเหตุการณ์ร้าย ถอยหลัง ${p.steps} ช่อง`,
            good_forward: (p) => `โชคดี เดินหน้าเพิ่ม ${p.steps} ช่อง`,
            warp_forward: (p) => `พลังพิเศษ เดินหน้า ${p.steps} ช่อง`,
            bonus_key: () => 'ลาภลอย! ได้รับกุญแจ 1 ชิ้น',
            bonus_armor: () => 'ลาภลอย! ได้รับเกราะศักดิ์สิทธิ์ 1 ชิ้น',
            bonus_warp: (p) => `ลาภลอย! เกิดการวาร์ป เดินหน้า ${p.steps} ช่อง`,
            secret_forward: (p) => `พบเหตุการณ์ลับ: ${I18N_JS.th.secretNames[p.secretIdx] || ''} พาเดินหน้า ${p.steps} ช่อง`,
            back_to_start: () => 'กลับมายังจุดเริ่มต้น',
            chain_loop_stop: (p) => `ตรวจสอบช่องใหม่: กลับมาช่อง ${p.pos} ที่เคยผ่านมาแล้ว หยุดการเดินทางเพื่อป้องกันการวนลูป`,
            chain_limit_stop: () => 'ตรวจสอบช่องใหม่: ครบขีดจำกัดเหตุการณ์ 3 ครั้งแล้ว จบการเคลื่อนที่สำหรับเทิร์นนี้ทันที',
            chain_continue: (p) => `ตรวจสอบช่องใหม่: ไปยังช่อง ${p.pos}`
        }
    },
    en: {
        defaultPlayerName: 'New Player',
        defaultHostName: 'Player 1',
        defaultJoinName: 'Joining Player',
        botPrefix: 'Bot ',
        botTag: '(Bot)',
        botNames: ['Bot Commander', 'Bot Spring', 'Bot Rainbow', 'Bot Zephyr', 'Bot Sage'],
        welcomeLobbyAnnounce: (name) => `Welcome ${name} to the Adventure Lobby`,
        manualOpened: 'Game guide opened',
        manualClosed: 'Game guide closed',
        roomItemLabel: (rId, pCount) => `<span><strong>Room ${rId}</strong> (${pCount}/6 players)</span>`,
        noRoomsWaiting: 'No rooms waiting for players...',
        roomCreatedAnnounce: (rId) => `Room created successfully. Room ID is ${rId}`,
        roomJoinedAnnounce: (rId) => `Joined room ${rId} successfully`,
        hostLeftAnnounce: 'The host left the game. Room closed.',
        playerDisconnectedBot: (name) => `${name} disconnected and was converted to a bot`,
        startGameAnnounce: 'The adventure has begun! Get ready!',
        lobbyRoomTitle: (rId) => `Game Room: ${rId}`,
        cellPrefix: (num) => `Space ${num}`,
        cellStart: ' Start Camp',
        cellFinish: ' Castle Finish',
        cellRest: ' Rest Camp, no event',
        cellTreasure: ' Treasure Chest',
        cellForward: ' Warp Boost',
        cellTrap: ' Pitfall Trap',
        cellWater: ' Rapids',
        cellGhost: ' Ghost Haunt',
        cellWarp: ' Cyclone',
        cellKey: ' Mystery Box',
        cellDoor: (dest) => ` Shortcut Door to Space ${dest}`,
        cellBonus: ' Windfall Bonus',
        cellSecret: ' Secret Event',
        turnStatusBar: (icon, name) => `Current Turn: ${icon} ${name}`,
        turnAnnounce: (name) => `It is ${name}'s turn`,
        allPlayersStatusHeader: 'All Players Status: ',
        playerStatusItem: (p) => `${p.name} Space ${p.pos} Keys ${p.keys} Armor ${p.armor}`,
        playerCardStats: (p) => `Space ${p.pos} 🔑 ${p.keys} 🛡️ ${p.armor}`,
        spaceTypes: {
            special: 'Special',
            rest: 'Rest Camp',
            treasure: 'Treasure Chest',
            forward: 'Warp Boost',
            trap: 'Pitfall Trap',
            water: 'Rapids',
            ghost: 'Ghost Haunt',
            warp: 'Cyclone',
            key: 'Mystery Box',
            door: 'Shortcut Door',
            bonus: 'Windfall Bonus',
            secret: 'Secret Event'
        },
        restNames: ['Warm Campfire', 'Peaceful Oasis', 'Forest Cabin', 'Giant Oak Tree', 'Traveler Pavilion'],
        doorNames: ['Dimensional Door', 'Time Gate', 'Ancient Stone Door', 'Magic Portal', 'Secret Door'],
        secretNames: ['Sacred Spirit', 'Mystic Guide Beast', 'Fairy of Light'],
        modalDoorTitle: (dName) => `Found ${dName}!`,
        modalDoorDesc: 'You have an Ancient Key. Use a key to open the shortcut door?',
        modalDoorConfirm: 'Use Key',
        modalDoorCancel: 'Skip',
        modalDangerTitle: 'Danger Ahead!',
        modalDangerDesc: (pos) => `Space ${pos} has a hazard! Use Sacred Armor to block it?`,
        modalDangerConfirm: 'Use Armor',
        modalDangerCancel: 'Skip',
        resultWinnerText: (icon, name) => `🎉 ${icon} ${name} reached Space 80 first!`,
        resultWinnerAnnounce: (name) => `The adventure has ended! The winner is ${name}!`,
        resultRankingsHeader: '<h3>Final Standings:</h3>',
        resultRankingRow: (name, pos) => `<p>${name} — Space ${pos}</p>`,
        events: {
            roll_dice: (p) => `${p.name} rolled a ${p.dice}`,
            walk_space: (p) => `${p.name} moved from Space ${p.from} to Space ${p.to}`,
            land_space: (p) => `${p.name} landed on Space ${p.pos}: ${I18N_JS.en.spaceTypes[p.spaceType] || p.spaceType}`,
            rest_arrive: (p) => `${p.name} arrived at Rest Camp (${I18N_JS.en.restNames[p.restIdx] || ''}) on Space ${p.pos}. Safe space, turn ends`,
            treasure_get: () => 'Found a Treasure Chest! Opened it and gained 1 Sacred Armor',
            treasure_empty: () => 'Found a Treasure Chest, but it was already opened—only dust inside',
            key_get: () => 'Found a Mystery Box! Opened it and gained 1 Ancient Key',
            key_empty: () => 'Found a Mystery Box, but it was already opened—only dust inside',
            door_open: (p) => `Used a key to unlock ${I18N_JS.en.doorNames[p.doorIdx] || 'the door'} and warped to Space ${p.dest}`,
            door_skip: () => 'Chose not to use a key. The door remains closed',
            door_nokey: (p) => `Found ${I18N_JS.en.doorNames[p.doorIdx] || 'a door'}, but have no keys to unlock it`,
            armor_block: () => 'Used Sacred Armor to block the hazard! Safe and sound!',
            armor_skip_back: (p) => `Chose not to use Sacred Armor and suffered the hazard, moving back ${p.steps} spaces`,
            bad_back: (p) => `Hazard triggered! Moved back ${p.steps} spaces`,
            good_forward: (p) => `Lucky break! Moved forward ${p.steps} spaces`,
            warp_forward: (p) => `Warp boost! Moved forward ${p.steps} spaces`,
            bonus_key: () => 'Windfall bonus! Received 1 Ancient Key',
            bonus_armor: () => 'Windfall bonus! Received 1 Sacred Armor',
            bonus_warp: (p) => `Windfall bonus! Warped forward ${p.steps} spaces`,
            secret_forward: (p) => `Secret Event: ${I18N_JS.en.secretNames[p.secretIdx] || ''} guided you forward ${p.steps} spaces`,
            back_to_start: () => 'Returned to the Start Camp',
            chain_loop_stop: (p) => `Checking new space: returned to previously visited Space ${p.pos}. Stopping movement to prevent a loop`,
            chain_limit_stop: () => 'Checking new space: reached the 3-event chain limit. Ending movement for this turn',
            chain_continue: (p) => `Checking new space: continuing to Space ${p.pos}`
        }
    }
};

function t() {
    return I18N_JS[getCurrentLang()] || I18N_JS.en;
}

function formatSyncedActionMessage(lastAction) {
    if (!lastAction) return '';
    const dict = t();
    if (lastAction.key && dict.events && typeof dict.events[lastAction.key] === 'function') {
        return dict.events[lastAction.key](lastAction.params || {});
    }
    return lastAction.msg || '';
}

window.onAdventureLanguageChanged = function() {
    if (gameState) {
        updateLobbyUI();
        if (gameState.status === 'playing') {
            renderBoardGrid();
            updateGameUI();
        } else if (gameState.status === 'ended' && document.getElementById('screen-result').classList.contains('active-screen')) {
            showResultScreen();
        }
    }
};

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
    myPlayerName = nameInput.value.trim() || t().defaultPlayerName;
    switchScreen('screen-lobby', 'lobby-heading');
    announceSR(t().welcomeLobbyAnnounce(myPlayerName));
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
        announceSR(t().manualOpened);
    } else {
        m.style.display = 'none';
        announceSR(t().manualClosed);
    }
};

// Initialize Room Listener for Main Menu
function initRoomListListener() {
    const roomsRef = ref(db, 'games/Adventure80/rooms');
    roomListListener = onValue(roomsRef, (snapshot) => {
        const rooms = snapshot.val() || {};
        const container = document.getElementById('room-list');
        if (!container || myPlayerId) return;

        const dict = t();
        container.innerHTML = '';
        let count = 0;
        for (const [rId, rData] of Object.entries(rooms)) {
            if (rData.status === 'waiting') {
                count++;
                const item = document.createElement('div');
                item.className = 'room-item';
                const pCount = rData.players ? Object.keys(rData.players).length : 0;
                item.innerHTML = dict.roomItemLabel(rId, pCount);
                item.onclick = () => window.joinAdventureRoom(rId);
                container.appendChild(item);
            }
        }
        if (count === 0) {
            container.innerHTML = `<div style="color:#a0a5b5; padding:10px;">${dict.noRoomsWaiting}</div>`;
        }
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
    if (!myPlayerName) myPlayerName = t().defaultHostName;

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
            announceSR(t().roomCreatedAnnounce(currentRoomId));
        }
    });
};

// Join Room Handler
window.joinAdventureRoom = function(rId) {
    if (!myPlayerName) myPlayerName = t().defaultJoinName;

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
                name: t().botPrefix + myPlayerName
            });
            
            document.getElementById('lobby-menu-section').style.display = 'none';
            document.getElementById('lobby-room-section').style.display = 'block';
            
            setupRoomListener();
            announceSR(t().roomJoinedAnnounce(currentRoomId));
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
        const botNames = t().botNames;
        const usedAnimals = Object.values(updatedPlayers).map(p => p.animal.icon);
        const availAnimals = ANIMAL_POOL.filter(a => !usedAnimals.includes(a.icon));

        updatedPlayers[botSlot] = {
            name: botNames[newBots - 1] || `${t().botPrefix}${newBots}`,
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
            announceSR(t().hostLeftAnnounce, 'assertive');
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
                    announceSR(t().playerDisconnectedBot(previousPlayersState[pId].name), 'polite');
                }
            }
            previousPlayersState = JSON.parse(JSON.stringify(gameState.players));
        }

        updateLobbyUI();

        // Synced Actions Announcements Listener & Audio sync
        if (gameState.lastAction && gameState.lastAction.ts > localLastActionTs) {
            localLastActionTs = gameState.lastAction.ts;
            const localizedMsg = formatSyncedActionMessage(gameState.lastAction);
            announceSR(localizedMsg, 'polite');
            if (gameState.lastAction.audioKeys && gameState.lastAction.audioKeys.length > 0) {
                playAudioSequence(gameState.lastAction.audioKeys);
            }
            
            // --- SYNC DICE VISUAL: อัปเดตตัวเลขผลลูกเต๋าให้ทุกอุปกรณ์เห็นตรงกันจากการทอยจริง ---
            if (gameState.lastAction.key === 'roll_dice' && gameState.lastAction.params && gameState.lastAction.params.dice) {
                const diceEl = document.getElementById('dice-visual');
                if (diceEl) {
                    diceEl.textContent = gameState.lastAction.params.dice;
                }
            } else if (gameState.lastAction.msg && gameState.lastAction.msg.includes('ทอยลูกเต๋าได้')) {
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
            VisualController.triggerFromAction(gameState.lastAction, pData);
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
    announceSR(t().startGameAnnounce, 'assertive');
    
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
    const dict = t();
    document.getElementById('lobby-room-title').textContent = dict.lobbyRoomTitle(currentRoomId);
    
    const pList = Object.values(gameState.players || {});
    document.getElementById('participant-count').textContent = `${pList.length}/6`;

    const listContainer = document.getElementById('participant-list');
    listContainer.innerHTML = '';
    pList.forEach(p => {
        const card = document.createElement('div');
        card.className = 'participant-card';
        card.innerHTML = `<span class="participant-avatar">${p.animal.icon}</span> <span>${p.name} ${p.isBot ? dict.botTag : ''}</span>`;
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
    const dict = t();

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
            let accessibleName = dict.cellPrefix(num);

            if (num === 1) {
                cell.classList.add('cell-start');
                iconHtml = '🏕️';
                accessibleName += dict.cellStart;
            } else if (num === 80) {
                cell.classList.add('cell-finish');
                iconHtml = '🏰';
                accessibleName += dict.cellFinish;
            } else if (config[num]) {
                const sp = config[num];
                cell.classList.add(`cell-${sp.type}`);
                
                if (sp.type === 'rest') { iconHtml = '⛺'; accessibleName += dict.cellRest; }
                else if (sp.type === 'treasure') { iconHtml = '📦'; accessibleName += dict.cellTreasure; }
                else if (sp.type === 'forward') { iconHtml = '🚀'; accessibleName += dict.cellForward; }
                else if (sp.type === 'trap') { iconHtml = '🕳️'; accessibleName += dict.cellTrap; }
                else if (sp.type === 'water') { iconHtml = '🌊'; accessibleName += dict.cellWater; }
                else if (sp.type === 'ghost') { iconHtml = '👻'; accessibleName += dict.cellGhost; }
                else if (sp.type === 'warp') { iconHtml = '🌀'; accessibleName += dict.cellWarp; }
                else if (sp.type === 'key') { iconHtml = '🔑'; accessibleName += dict.cellKey; }
                else if (sp.type === 'door') { iconHtml = '🚪'; accessibleName += dict.cellDoor(sp.dest); }
                else if (sp.type === 'bonus') { iconHtml = '🎁'; accessibleName += dict.cellBonus; }
                else if (sp.type === 'secret') { iconHtml = '🔮'; accessibleName += dict.cellSecret; }
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
    const dict = t();

    const pOrder = gameState.playerOrder || Object.keys(gameState.players);
    const playersArr = pOrder.map(pId => [pId, gameState.players[pId]]);
    const currentTurnKey = playersArr[gameState.turnIndex][0];
    const currentTurnPlayer = playersArr[gameState.turnIndex][1];

    document.getElementById('game-status-bar').textContent = dict.turnStatusBar(currentTurnPlayer.animal.icon, currentTurnPlayer.name);

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
            const turnMsg = dict.turnAnnounce(currentTurnPlayer.name);
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
    let allStatusText = dict.allPlayersStatusHeader;
    playersArr.forEach(([pId, p], index) => {
        allStatusText += dict.playerStatusItem(p);
        if (index < playersArr.length - 1) allStatusText += ", ";
    });
    
    cardsContainer.setAttribute('aria-label', allStatusText);
    cardsContainer.setAttribute('tabindex', '0');
    cardsContainer.setAttribute('role', 'group');

    // Render visual cards for sighted users and hide them from screen reader to prevent duplicate reading
    playersArr.forEach(([pId, p]) => {
        const card = document.createElement('div');
        card.className = `p-status-card ${pId === currentTurnKey ? 'active-turn' : ''}`;
        card.setAttribute('aria-hidden', 'true');
        card.innerHTML = `
            <div><strong>${p.animal.icon} ${p.name}</strong></div>
            <div>${dict.playerCardStats(p)}</div>
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

// Emits an action with structured i18n key + params so all clients render in their own language
async function syncActionEmit(key, params = {}, audioKeys = []) {
    const ts = Date.now() + Math.random();
    const fallbackMsg = (I18N_JS.th.events[key] ? I18N_JS.th.events[key](params) : key);
    await update(ref(db), {
        [`games/Adventure80/rooms/${currentRoomId}/lastAction`]: {
            key,
            params,
            msg: fallbackMsg,
            ts,
            audioKeys
        }
    });
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
    
    await syncActionEmit('roll_dice', { name: pData.name, dice: diceRoll, pos: pData.pos }, ['dice.mp3']);
    
    let currentPos = pData.pos;
    let targetPos = currentPos + diceRoll;
    if (targetPos > 80) targetPos = 80 - (targetPos - 80); // Bounce back rule

    await syncActionEmit('walk_space', { name: pData.name, from: currentPos, to: targetPos, pos: targetPos }, ['walk.mp3']);
    
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

        let landingAudio = null;
        if (sp.type === 'rest') { landingAudio = 'sabuy.mp3'; }
        else if (sp.type === 'treasure') { landingAudio = 'box1.mp3'; }
        else if (sp.type === 'forward') { landingAudio = 'forward.mp3'; }
        else if (sp.type === 'trap') { landingAudio = 'hole.mp3'; }
        else if (sp.type === 'water') { landingAudio = 'water.mp3'; }
        else if (sp.type === 'ghost') { landingAudio = 'ghost.mp3'; }
        else if (sp.type === 'warp') { landingAudio = 'cyclone.mp3'; }
        else if (sp.type === 'key') { landingAudio = 'box2.mp3'; }
        else if (sp.type === 'door') { landingAudio = 'door1.mp3'; }
        else if (sp.type === 'bonus') { landingAudio = 'box4.mp3'; }
        else if (sp.type === 'secret') { landingAudio = 'box3.mp3'; }

        await syncActionEmit('land_space', { name: pData.name, pos: currentPos, spaceType: sp.type }, landingAudio ? [landingAudio] : []);

        // Handle Effect
        if (sp.type === 'rest') {
            const restIdx = Math.floor(Math.random() * t().restNames.length);
            await syncActionEmit('rest_arrive', { name: pData.name, restIdx, pos: currentPos });
            break; // Rule 5: Ends event chain immediately
        }
        
        let movedToNewSpace = false;

        if (sp.type === 'treasure') {
            if (sp.charges > 0) {
                sp.charges--;
                await update(ref(db), { [`games/Adventure80/rooms/${currentRoomId}/boardConfig/${currentPos}/charges`]: sp.charges });
                pData.armor++;
                await syncStateDB(pId, { armor: pData.armor });
                await syncActionEmit('treasure_get', { pos: currentPos }, ['akey3.mp3']);
            } else {
                await syncActionEmit('treasure_empty', { pos: currentPos }, ['akey2.mp3']);
            }
            break; // Event chain ends as player did not change location
        } else if (sp.type === 'key') {
            if (sp.charges > 0) {
                sp.charges--;
                await update(ref(db), { [`games/Adventure80/rooms/${currentRoomId}/boardConfig/${currentPos}/charges`]: sp.charges });
                pData.keys++;
                await syncStateDB(pId, { keys: pData.keys });
                await syncActionEmit('key_get', { pos: currentPos }, ['akey1.mp3']);
            } else {
                await syncActionEmit('key_empty', { pos: currentPos }, ['akey2.mp3']);
            }
            break; // Does not change location
        } else if (sp.type === 'door') {
            const dict = t();
            const doorIdx = Math.floor(Math.random() * dict.doorNames.length);
            const dName = dict.doorNames[doorIdx];
            if (pData.keys > 0) {
                let useKey = true;
                if (!autoMode && pId === myPlayerId) {
                    useKey = await showModalAsync(dict.modalDoorTitle(dName), dict.modalDoorDesc, dict.modalDoorConfirm, dict.modalDoorCancel);
                }
                if (useKey) {
                    pData.keys--;
                    await syncStateDB(pId, { keys: pData.keys });
                    await syncActionEmit('door_open', { doorIdx, dest: sp.dest, pos: sp.dest }, ['door2.mp3', 'walk.mp3']);
                    currentPos = sp.dest;
                    movedToNewSpace = true;
                } else {
                    await syncActionEmit('door_skip', { pos: currentPos }, ['shieldno.mp3']);
                    break;
                }
            } else {
                await syncActionEmit('door_nokey', { doorIdx, pos: currentPos });
                break;
            }
        } else if (['trap', 'water', 'ghost', 'warp'].includes(sp.type)) {
            const dict = t();
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
                    useArmor = await showModalAsync(dict.modalDangerTitle, dict.modalDangerDesc(currentPos), dict.modalDangerConfirm, dict.modalDangerCancel);
                }
                if (useArmor) {
                    pData.armor--;
                    await syncStateDB(pId, { armor: pData.armor });
                    await syncActionEmit('armor_block', { pos: currentPos }, ['shield.mp3']);
                    break;
                } else {
                    const nextPos = Math.max(1, currentPos + randVal);
                    await syncActionEmit('armor_skip_back', { steps: Math.abs(randVal), pos: nextPos }, ['shieldno.mp3', 'walk.mp3']);
                    currentPos = nextPos;
                    movedToNewSpace = true;
                }
            } else if (isBad) {
                const nextPos = Math.max(1, currentPos + randVal);
                await syncActionEmit('bad_back', { steps: Math.abs(randVal), pos: nextPos }, ['walk.mp3']);
                currentPos = nextPos;
                movedToNewSpace = true;
            } else {
                let targetPos = currentPos + randVal;
                if (targetPos > 80) targetPos = 80 - (targetPos - 80);
                await syncActionEmit('good_forward', { steps: randVal, pos: targetPos }, ['forward.mp3', 'walk.mp3']);
                currentPos = targetPos;
                movedToNewSpace = true;
            }
        } else if (sp.type === 'forward') {
            let randVal = Math.floor(Math.random() * 10) + 3;
            let targetPos = currentPos + randVal;
            if (targetPos > 80) targetPos = 80 - (targetPos - 80);
            await syncActionEmit('warp_forward', { steps: randVal, pos: targetPos }, ['forward.mp3', 'walk.mp3']);
            currentPos = targetPos;
            movedToNewSpace = true;
        } else if (sp.type === 'bonus') {
            const randEffect = Math.floor(Math.random() * 3);
            if (randEffect === 0) {
                pData.keys++;
                await syncStateDB(pId, { keys: pData.keys });
                await syncActionEmit('bonus_key', { pos: currentPos }, ['akey1.mp3']);
                break;
            } else if (randEffect === 1) {
                pData.armor++;
                await syncStateDB(pId, { armor: pData.armor });
                await syncActionEmit('bonus_armor', { pos: currentPos }, ['akey3.mp3']);
                break;
            } else {
                let randVal = Math.floor(Math.random() * 10) + 3;
                let targetPos = currentPos + randVal;
                if (targetPos > 80) targetPos = 80 - (targetPos - 80);
                await syncActionEmit('bonus_warp', { steps: randVal, pos: targetPos }, ['forward.mp3', 'walk.mp3']);
                currentPos = targetPos;
                movedToNewSpace = true;
            }
        } else if (sp.type === 'secret') {
            const secretIdx = Math.floor(Math.random() * t().secretNames.length);
            let randVal = Math.floor(Math.random() * 10) + 3;
            let targetPos = currentPos + randVal;
            if (targetPos > 80) targetPos = 80 - (targetPos - 80);
            await syncActionEmit('secret_forward', { secretIdx, steps: randVal, pos: targetPos }, ['forward.mp3', 'walk.mp3']);
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
                await syncActionEmit('back_to_start', { pos: 1 }, ['sabuy.mp3']);
            }

            if (visited.has(currentPos)) {
                await syncActionEmit('chain_loop_stop', { pos: currentPos });
                break;
            }
            
            visited.add(currentPos);
            chainCount++;
            
            if (chainCount >= 3) {
                await syncActionEmit('chain_limit_stop', { pos: currentPos });
                break; // Stop Chain
            } else {
                await syncActionEmit('chain_continue', { pos: currentPos });
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
    const dict = t();
    
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

    document.getElementById('result-winner-text').textContent = dict.resultWinnerText(winner.animal.icon, winner.name);
    announceSR(dict.resultWinnerAnnounce(winner.name), 'assertive');

    const rankingsBox = document.getElementById('result-rankings');
    rankingsBox.innerHTML = dict.resultRankingsHeader;

    const sorted = Object.values(gameState.players).sort((a, b) => b.pos - a.pos);
    sorted.forEach((p) => {
        rankingsBox.innerHTML += dict.resultRankingRow(p.name, p.pos);
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
                name: t().botPrefix + myPlayerName
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

    triggerFromAction: function(action, pData) {
        if (this.isReducedMotion() || !action) return;
        const params = action.params || {};
        const key = action.key || '';
        const msg = action.msg || '';

        let targetPos = params.pos || (pData ? pData.pos : 1);
        if (!params.pos && msg) {
            const match = msg.match(/ช่อง\s*(\d+)/);
            if (match) targetPos = parseInt(match[1]);
        }

        const cell = document.getElementById(`cell-${targetPos}`);
        if (!cell) return;
        
        let activePawn = null;
        const pawnHolder = document.getElementById(`pawns-holder-${targetPos}`);
        if (pawnHolder && pData) {
            Array.from(pawnHolder.querySelectorAll('.pawn-token')).forEach(p => {
                if (p.textContent.includes(pData.name)) activePawn = p;
            });
        }

        // Event Matching Visuals (supports structured key + legacy msg fallback)
        if (key === 'roll_dice' || msg.includes('ทอยลูกเต๋าได้')) {
            this.animateDice();
        } else if (
            ['walk_space', 'armor_skip_back', 'bad_back', 'good_forward', 'warp_forward', 'secret_forward', 'back_to_start'].includes(key) ||
            msg.includes('เดินจากช่อง') || msg.includes('ถอยหลัง') || msg.includes('เดินหน้า') || msg.includes('กลับมายังจุดเริ่มต้น')
        ) {
            if (key === 'armor_skip_back' || key === 'bad_back' || msg.includes('ถอยหลัง')) {
                this.createEffect(cell, '⬅️', 'effect-backward', 1500);
            } else if (['good_forward', 'warp_forward', 'secret_forward'].includes(key) || msg.includes('เดินหน้า') || msg.includes('พลังพิเศษ') || msg.includes('พาเดินหน้า')) {
                this.createEffect(cell, '➡️', 'effect-forward', 1500);
            } else {
                this.createEffect(cell, '👣', 'effect-walk', 1000);
            }
            if (activePawn) {
                activePawn.classList.add('visual-bounce');
                setTimeout(() => activePawn.classList.remove('visual-bounce'), 500);
            }
        } else if ((key === 'land_space' && params.spaceType === 'water') || msg.includes('น้ำเชี่ยว')) {
            this.createEffect(cell, '🌊', 'effect-water', 2000);
            if (activePawn) { activePawn.classList.add('visual-scale'); setTimeout(() => activePawn.classList.remove('visual-scale'), 1000); }
        } else if ((key === 'land_space' && params.spaceType === 'warp') || msg.includes('ไซโคลน')) {
            this.createEffect(cell, '🌪️', 'effect-cyclone', 2000);
            if (activePawn) { activePawn.classList.add('visual-rotate'); setTimeout(() => activePawn.classList.remove('visual-rotate'), 1000); }
        } else if ((key === 'land_space' && params.spaceType === 'trap') || msg.includes('หลุมพราง') || msg.includes('เหตุการณ์ร้าย')) {
            this.createEffect(cell, '🕳️', 'effect-trap', 2000);
            if (activePawn) { activePawn.classList.add('visual-drop'); setTimeout(() => activePawn.classList.remove('visual-drop'), 1000); }
        } else if ((key === 'land_space' && params.spaceType === 'ghost') || msg.includes('ผีหลอก')) {
            this.createEffect(cell, '👻', 'effect-ghost', 2000);
            if (activePawn) { activePawn.classList.add('visual-shake'); setTimeout(() => activePawn.classList.remove('visual-shake'), 1000); }
        } else if (key === 'bonus_warp' || msg.includes('วาร์ปไปยังช่อง')) {
            this.createEffect(cell, '🌀', 'effect-warp', 2000);
        } else if ((key === 'land_space' && params.spaceType === 'secret') || msg.includes('เหตุการณ์ลับ')) {
            this.createEffect(cell, '✨', 'effect-secret', 2000);
        } else if (['bonus_key', 'bonus_armor'].includes(key) || (key === 'land_space' && params.spaceType === 'bonus') || msg.includes('ลาภลอย')) {
            this.createEffect(cell, '✨✨✨', 'effect-bonus', 2000);
        } else if (['treasure_get', 'treasure_empty'].includes(key) || msg.includes('หีบสมบัติ')) {
            if (key === 'treasure_empty' || msg.includes('เหลือเพียงเศษฝุ่น')) this.createEffect(cell, '💨', 'effect-dust', 1500);
            else this.createEffect(cell, '🛡️✨', 'effect-armor-fly', 2000);
        } else if (['key_get', 'key_empty'].includes(key) || msg.includes('กล่องลึกลับ')) {
            if (key === 'key_empty' || msg.includes('เหลือเพียงเศษฝุ่น')) this.createEffect(cell, '💨', 'effect-dust', 1500);
            else this.createEffect(cell, '🔑✨', 'effect-key-fly', 2000);
        } else if (['door_open', 'door_skip', 'door_nokey'].includes(key) || msg.includes('ประตู')) {
            if (key === 'door_open' || msg.includes('ใช้กุญแจเปิด')) this.createEffect(cell, '🔑✨🚪', 'effect-door-open', 2000);
            else if (['door_skip', 'door_nokey'].includes(key) || msg.includes('ไม่มีกุญแจ') || msg.includes('ประตูยังคงปิด')) this.createEffect(cell, '🚪❌', 'effect-door-closed', 1500);
        } else if (key === 'rest_arrive' || msg.includes('จุดพักผ่อน')) {
            this.createEffect(cell, '🔥', 'effect-rest', 2500);
        } else if (key === 'armor_block' || msg.includes('ใช้เกราะศักดิ์สิทธิ์ป้องกันผลเสียสำเร็จ')) {
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
                const dict = t();
                let allStatusText = dict.allPlayersStatusHeader;
                const pOrder = gameState.playerOrder || Object.keys(gameState.players);
                pOrder.forEach((pId, index) => {
                    const p = gameState.players[pId];
                    allStatusText += dict.playerStatusItem(p);
                    if (index < pOrder.length - 1) allStatusText += ", ";
                });
                announceSR(allStatusText, 'polite');
            }
        }
    });
});