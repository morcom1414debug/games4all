import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getDatabase, ref, set, onValue, remove, onDisconnect, runTransaction } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-database.js";

// Global Utility & Guard States
window.joinOrder = []; 
window.isStartingRound = false; 
window.hasAnnouncedMyColor = false; 

// =========================================================
// [พื้นที่สำหรับพัฒนาต่อ #1] ระบบตรวจสอบกฎพิเศษ (Special Win)
// =========================================================
let currentSpecialWinType = null;
window.submitSpecialWin = () => {
    if (currentSpecialWinType) clientAction('SPECIAL_WIN', currentSpecialWinType);
};

function getSpecialWinType(hand, player) {
    if (!hand || hand.length === 0) return null;

    // 1. ตรวจ 3A ก่อนเสมอ (ใช้ได้ตั้งแต่ PRE_GAME และตลอดทั้งเกม)
    const aceCount = hand.filter(c => c.rank === 'A').length;
    if (aceCount >= 3) return '3A';

    // 2. ถ้าไม่ใช่ Turn แรกของผู้เล่นคนนี้ (หมดรอบแรกของตัวเองไปแล้ว)
    // จะหมดสิทธิ์ใช้กฎพิเศษกลุ่มที่ 2 ทันที
    if (player && player.hasFinishedFirstTurn) {
        return null; 
    }

    // 3. กฎพิเศษอื่นๆ (เฉพาะ Turn แรกของผู้เล่นแต่ละคนเท่านั้น)
    const counts = {};
    let totalPoints = 0;
    const suits = new Set();
    const ranks = [];

    hand.forEach(c => {
        counts[c.rank] = (counts[c.rank] || 0) + 1;
        suits.add(c.suit);
        
        // คำนวณแต้มสำหรับกฎ 50
        let val = 0;
        if (['J', 'Q', 'K'].includes(c.rank)) val = 10;
        else if (c.rank === 'A') val = 1;
        else val = parseInt(c.rank);
        totalPoints += val;

        // สำหรับกฎเรียง
        const rankOrder = ['A','2','3','4','5','6','7','8','9','10','J','Q','K'];
        ranks.push(rankOrder.indexOf(c.rank));
    });

    // ตอง (ไพ่ 3 ใบเหมือนกัน และต้องมีไพ่ในมือ 5 ใบ)
    const hasThreeOfAKind = Object.values(counts).some(count => count >= 3);
    if (hasThreeOfAKind && hand.length === 5) return 'ตอง';

    // 50 แต้ม (แต้มรวม 50 พอดี และมีไพ่ 5 ใบ)
    if (totalPoints === 50 && hand.length === 5) return '50';

    // สี (ดอกเดียวกันทั้งหมด 5 ใบ)
    if (suits.size === 1 && hand.length === 5) return 'สี';

    // 3 คู่ (คู่ 3 คู่ และต้องมีไพ่ 6 ใบ ซึ่งหมายถึงรับแจก 5 จั่ว 1 ในเทิร์นแรก)
    const pairsCount = Object.values(counts).filter(count => count >= 2).length;
    if (pairsCount >= 3) return '3คู่';

    // เรียง (ไพ่ 5 ใบ ลำดับติดกัน)
    if (hand.length === 5) {
        ranks.sort((a, b) => a - b);
        let isStraight = true;
        for (let i = 0; i < ranks.length - 1; i++) {
            if (ranks[i + 1] - ranks[i] !== 1) {
                isStraight = false;
                break;
            }
        }
        if (isStraight) return 'เรียง';
    }

    return null; // ไม่มีกฎพิเศษที่ตรงเงื่อนไข
}
// =========================================================

const delay = ms => new Promise(res => setTimeout(res, ms));

// =========================================================
// Audio Setup (คงไว้ 100% เพิ่มไฟล์เสียงใหม่ และระบบจัดการคิว/AudioContext ส่วนกลาง)
// =========================================================
let audioCtx = null;
let audioResumePromise = null;
const audioBuffers = {};
const soundLoadStatus = {}; // 'idle' | 'loading' | 'loaded' | 'failed'
const soundLoadPromises = {};
let bgmSource = null;
let isBgmWanted = false;

const pendingSoundQueue = [];
let isProcessingSoundQueue = false;
let soundEventSeq = 0;
let lastHandledRemoteSoundId = null;
const recentPlayedSignatureMap = new Map();
const pendingPeerSoundPayloads = new Map();

const soundFiles = {
    '1': 'audio/1.mp3', 'select': 'audio/select.mp3', 'start': 'audio/start.mp3',
    'bgm': 'audio/bgm.mp3', 'jua': 'audio/jua.mp3', 'turn': 'audio/turn.mp3',
    'follow': 'audio/follow.mp3', 'cang': 'audio/cang.mp3', 'cang25': 'audio/cang25.mp3',
    'win': 'audio/win.mp3', 'lost': 'audio/lost.mp3', '60': 'audio/60.mp3', 'no': 'audio/no.mp3',
    'knock': 'audio/knock.mp3', 'special': 'audio/special.mp3' // เพิ่มเสียง knock และ special
};

function getOrCreateAudioContext() {
    if (!audioCtx) {
        const AudioContextClass = window.AudioContext || window.webkitAudioContext;
        if (AudioContextClass) {
            try {
                audioCtx = new AudioContextClass();
                audioCtx.onstatechange = () => {
                    if (audioCtx && audioCtx.state === 'running') {
                        processPendingSoundQueue();
                        ensureBgmPlayingIfWanted();
                    }
                };
            } catch (e) {
                console.warn("Could not create AudioContext", e);
            }
        }
    }
    return audioCtx;
}

getOrCreateAudioContext();

function ensureAudioContextReady() {
    const ctx = getOrCreateAudioContext();
    if (!ctx) return Promise.resolve(false);
    if (ctx.state === 'running') return Promise.resolve(true);

    if (!audioResumePromise) {
        try {
            const p = ctx.resume();
            if (p && typeof p.then === 'function') {
                audioResumePromise = p.then(() => {
                    audioResumePromise = null;
                    return ctx.state === 'running';
                }).catch(() => {
                    audioResumePromise = null;
                    return ctx.state === 'running';
                });
            } else {
                audioResumePromise = Promise.resolve(ctx.state === 'running');
                audioResumePromise = null;
            }
        } catch (e) {
            audioResumePromise = null;
            return Promise.resolve(false);
        }
    }
    return audioResumePromise || Promise.resolve(ctx.state === 'running');
}

function decodeAudioDataAsync(ctx, arrayBuffer) {
    return new Promise((resolve, reject) => {
        try {
            const res = ctx.decodeAudioData(
                arrayBuffer,
                (decoded) => resolve(decoded),
                (err) => reject(err)
            );
            if (res && typeof res.then === 'function') {
                res.then(resolve).catch(reject);
            }
        } catch (e) {
            reject(e);
        }
    });
}

function loadSingleSound(key) {
    if (!soundFiles[key]) {
        soundLoadStatus[key] = 'failed';
        return Promise.resolve(null);
    }
    if (soundLoadStatus[key] === 'loaded' && audioBuffers[key]) {
        return Promise.resolve(audioBuffers[key]);
    }
    if (soundLoadStatus[key] === 'loading' && soundLoadPromises[key]) {
        return soundLoadPromises[key];
    }

    soundLoadStatus[key] = 'loading';
    soundLoadPromises[key] = (async () => {
        try {
            const response = await fetch(soundFiles[key]);
            if (!response.ok) throw new Error(`HTTP status ${response.status}`);
            const arrayBuffer = await response.arrayBuffer();
            const ctx = getOrCreateAudioContext();
            if (!ctx) throw new Error("AudioContext unavailable");
            const audioBuffer = await decodeAudioDataAsync(ctx, arrayBuffer);
            audioBuffers[key] = audioBuffer;
            soundLoadStatus[key] = 'loaded';
            return audioBuffer;
        } catch (e) {
            soundLoadStatus[key] = 'failed';
            console.warn("Could not load sound", key, e);
            return null;
        } finally {
            delete soundLoadPromises[key];
            processPendingSoundQueue();
            if (key === 'bgm') ensureBgmPlayingIfWanted();
        }
    })();

    return soundLoadPromises[key];
}

async function loadSounds() {
    const keys = Object.keys(soundFiles);
    await Promise.allSettled(keys.map(key => loadSingleSound(key)));
}
loadSounds();

function startBufferSourceNow(key, loop = false) {
    const ctx = getOrCreateAudioContext();
    if (!ctx || !audioBuffers[key] || ctx.state !== 'running') return null;
    try {
        const source = ctx.createBufferSource();
        source.buffer = audioBuffers[key];
        source.connect(ctx.destination);
        source.loop = loop;
        source.start(0);
        return source;
    } catch (e) {
        console.warn("Error starting sound source", key, e);
        return null;
    }
}

function ensureBgmPlayingIfWanted() {
    if (!isBgmWanted) return;
    if (bgmSource) return;
    const ctx = getOrCreateAudioContext();
    if (!ctx || ctx.state !== 'running') return;
    if (soundLoadStatus['bgm'] === 'loaded' && audioBuffers['bgm']) {
        bgmSource = startBufferSourceNow('bgm', true);
    } else if (soundLoadStatus['bgm'] !== 'loading' && soundLoadStatus['bgm'] !== 'failed') {
        loadSingleSound('bgm');
    }
}

async function processPendingSoundQueue() {
    if (isProcessingSoundQueue) return;
    if (pendingSoundQueue.length === 0) return;

    const ctx = getOrCreateAudioContext();
    if (!ctx) return;

    isProcessingSoundQueue = true;
    try {
        if (ctx.state !== 'running') {
            const resumed = await ensureAudioContextReady();
            if (!resumed && ctx.state !== 'running') {
                return;
            }
        }

        while (pendingSoundQueue.length > 0) {
            if (ctx.state !== 'running') break;

            const item = pendingSoundQueue[0];
            const { key, loop } = item;

            if (!soundFiles[key] || soundLoadStatus[key] === 'failed') {
                pendingSoundQueue.shift();
                continue;
            }

            if (soundLoadStatus[key] !== 'loaded' || !audioBuffers[key]) {
                const buffer = await loadSingleSound(key);
                if (!buffer) {
                    pendingSoundQueue.shift();
                    continue;
                }
                if (ctx.state !== 'running') break;
            }

            pendingSoundQueue.shift();
            if (key === 'bgm' && loop) {
                if (isBgmWanted) {
                    if (bgmSource) {
                        try { bgmSource.stop(); } catch (e) {}
                        bgmSource = null;
                    }
                    bgmSource = startBufferSourceNow('bgm', true);
                }
            } else {
                startBufferSourceNow(key, loop);
            }
        }
    } finally {
        isProcessingSoundQueue = false;
    }
}

function playSound(key, loop = false) {
    if (!key || !soundFiles[key]) return null;

    if (!loop) {
        const now = Date.now();
        const lastTime = recentPlayedSignatureMap.get(key) || 0;
        if (now - lastTime < 35) {
            return null;
        }
        recentPlayedSignatureMap.set(key, now);
    }

    const ctx = getOrCreateAudioContext();
    if (!ctx) return null;

    const isBufferReady = (soundLoadStatus[key] === 'loaded' && !!audioBuffers[key]);
    const isContextRunning = (ctx.state === 'running');

    if (isContextRunning && isBufferReady) {
        return startBufferSourceNow(key, loop);
    }

    if (soundLoadStatus[key] !== 'loaded' && soundLoadStatus[key] !== 'loading') {
        loadSingleSound(key);
    }

    if (key === 'bgm' && loop) {
        const hasBgmQueued = pendingSoundQueue.some(item => item.key === 'bgm' && item.loop);
        if (!hasBgmQueued) {
            pendingSoundQueue.push({ key, loop: true, timestamp: Date.now() });
        }
    } else {
        pendingSoundQueue.push({ key, loop, timestamp: Date.now() });
    }

    ensureAudioContextReady().then(() => {
        processPendingSoundQueue();
        ensureBgmPlayingIfWanted();
    });

    return null;
}

function stopBGM() {
    isBgmWanted = false;
    for (let i = pendingSoundQueue.length - 1; i >= 0; i--) {
        if (pendingSoundQueue[i].key === 'bgm' && pendingSoundQueue[i].loop) {
            pendingSoundQueue.splice(i, 1);
        }
    }
    if (bgmSource) {
        try { bgmSource.stop(); } catch(e) {}
        bgmSource = null;
    }
}

function playBGM() {
    stopBGM();
    isBgmWanted = true;
    bgmSource = playSound('bgm', true);
}

function unlockAudioOnUserGesture() {
    const ctx = getOrCreateAudioContext();
    if (!ctx) return;
    if (ctx.state !== 'running') {
        ensureAudioContextReady().then((ready) => {
            if (ready) {
                processPendingSoundQueue();
                ensureBgmPlayingIfWanted();
            }
        });
    } else if (pendingSoundQueue.length > 0 || (isBgmWanted && !bgmSource)) {
        processPendingSoundQueue();
        ensureBgmPlayingIfWanted();
    }
}

['click', 'touchstart', 'keydown', 'pointerdown'].forEach(evtName => {
    document.addEventListener(evtName, unlockAudioOnUserGesture, { passive: true, once: false });
});

function sendSoundMessageToConn(conn, payload) {
    if (!conn) return;
    if (conn.open) {
        try {
            conn.send(payload);
        } catch (e) {
            console.warn("Failed to send sound message to peer", conn.peer, e);
        }
    } else {
        if (!pendingPeerSoundPayloads.has(conn.peer)) {
            pendingPeerSoundPayloads.set(conn.peer, []);
            conn.on('open', () => {
                const queued = pendingPeerSoundPayloads.get(conn.peer) || [];
                pendingPeerSoundPayloads.delete(conn.peer);
                queued.forEach(msg => {
                    try {
                        if (conn.open) conn.send(msg);
                    } catch (e) {}
                });
            });
        }
        pendingPeerSoundPayloads.get(conn.peer).push(payload);
    }
}

function handleIncomingPlaySoundData(data) {
    if (!data || !data.sound) return;
    if (isHost) return; // ป้องกันเสียงซ้ำในเครื่อง Host กรณีได้รับ Broadcast ย้อนกลับ
    if (data.soundEventId) {
        if (lastHandledRemoteSoundId === data.soundEventId) return;
        lastHandledRemoteSoundId = data.soundEventId;
    }
    playSound(data.sound);
}

function triggerSound(soundKey) {
    if (!soundKey) return;
    const soundEventId = `${myPeerId || 'host'}_${Date.now()}_${++soundEventSeq}`;
    if (myPeerId) playSound(soundKey);
    const payload = { type: 'PLAY_SOUND', sound: soundKey, soundEventId };
    guestConnections.forEach(c => {
        if (c && c.peer !== myPeerId) {
            sendSoundMessageToConn(c, payload);
        }
    });
}

// =========================================================
// Firebase Config
// =========================================================
const firebaseConfig = {
    apiKey: "AIzaSyDvcdgsyT5sDdYTYKIqetzNL9Be-MFC0l4",
    authDomain: "xo-game-134ec.firebaseapp.com",
    databaseURL: "https://xo-game-134ec-default-rtdb.asia-southeast1.firebasedatabase.app",
    projectId: "xo-game-134ec",
    storageBucket: "xo-game-134ec.firebasestorage.app",
    messagingSenderId: "318375224157",
    appId: "1:318375224157:web:e88070bc0795859cb77eb4"
};
const app = initializeApp(firebaseConfig);
const db = getDatabase(app);
const GAME_DB_ID = "Cang"; 

const COLORS = [
    { id: 'red', hex: 'var(--red)', name: 'แดง' }, { id: 'yellow', hex: 'var(--yellow)', name: 'เหลือง' },
    { id: 'green', hex: 'var(--green)', name: 'เขียว' }, { id: 'blue', hex: 'var(--blue)', name: 'ฟ้า' },
    { id: 'pink', hex: 'var(--pink)', name: 'ชมพู' }, { id: 'orange', hex: 'var(--orange)', name: 'ส้ม' }
];
const THAI_SUITS = { '♠': 'โพดำ', '♥': 'โพแดง', '♦': 'ข้าวหลามตัด', '♣': 'ดอกจิก' };
const THAI_RANKS = { 'A': 'เอซ', '2': 'สอง', '3': 'สาม', '4': 'สี่', '5': 'ห้า', '6': 'หก', '7': 'เจ็ด', '8': 'แปด', '9': 'เก้า', '10': 'สิบ', 'J': 'แจ็ค', 'Q': 'แหม่ม', 'K': 'คิง' };

const BOT_NAMES_POOL = [
    "บอทเมเปิ้ล",
    "บอทตาหวาน",
    "บอทข้าวปั้น",
    "บอทโมจิ",
    "บอทถุงเงิน",
    "บอทมะลิ",
    "บอทไข่มุก",
    "บอทสายไหม",
    "บอทโกโก้",
    "บอทเจ้าถั่ว",
    "บอทใบชา",
    "บอทปุยนุ่น",
    "บอทน้ำผึ้ง",
    "บอทคุกกี้",
    "บอทลูกพลับ",
    "บอทข้าวเหนียว",
    "บอทฟองนม",
    "บอทปลาทู",
    "บอทจุ๊บจิ๊บ",
    "บอทเจ้าขนม"
];

let peer = null, myPeerId = null, roomId = null, isHost = false;
let hostConnection = null, guestConnections = []; 
let roomPlayers = [], botColors = [], botNames = [], currentBotCount = 0, selectedColorId = 'red';
let myPlayerName = "", isNameConfirmed = false;
let deck = [], discardPile = [];
let gameState = { turnIndex: 0, players: [], status: 'WAITING', topCardOwnerId: null, flowSourceId: null, skipPreVotes: [] };
let localPlayerState = { hand: [], hasDrawn: false, hasDiscarded: false, discardedRank: null, points: 200, turnCount: 0, hasFinishedFirstTurn: false, hasFlowedThisTurn: false, temporaryBot: false, returnRequested: false };
let globalPlayersMap = {}; 
let turnTimerInterval = null;
let announcerTimeout = null;
let hostHeartbeatTimer = null;
let heartbeatInterval = null;

// =========================================================
// Timeout, Temporary Bot & Disconnect States
// =========================================================
let currentTurnMaxTime = 40;
let isProcessingTurnTimeout = false;
let isEndingTurn = false;
let activeBotSequencePlayerId = null;
let botResponseTimers = {}; // Host map: { [playerId]: { interval, timeLeft, responded } }
let localBotModalTimer = null;
let hasSubmittedBotModalResponse = false;
let isLocalTempBotModalOpen = false;
let processedDisconnects = new Set();
let isClosingHostRoom = false;

function randomizeAllBotNames(count) {
    let shuffled = [...BOT_NAMES_POOL].sort(() => Math.random() - 0.5);
    botNames = shuffled.slice(0, count);
}

function triggerHostDisconnect() {
    clearTimeout(hostHeartbeatTimer);
    if (localBotModalTimer) {
        clearInterval(localBotModalTimer);
        localBotModalTimer = null;
    }
    closeTemporaryBotModal();

    const nameScreen = document.getElementById('name-entry-screen');
    if (nameScreen) nameScreen.style.display = 'none';
    document.getElementById('setup-screen').style.display = 'none';
    document.getElementById('waiting-room-screen').style.display = 'none';
    document.getElementById('game-screen').style.display = 'none';
    document.getElementById('result-modal').style.display = 'none';
    document.getElementById('countdown-screen').style.display = 'none';
    document.getElementById('kang-animation-screen').style.display = 'none';
    document.getElementById('transition-overlay').style.display = 'none';
    
    const dcScreen = document.getElementById('disconnect-screen');
    dcScreen.style.display = 'flex';
    setTimeout(() => {
        const hd = document.getElementById('disconnect-heading');
        if(hd) hd.focus();
    }, 100);
}

function resetHeartbeat() {
    if(isHost) return; 
    clearTimeout(hostHeartbeatTimer);
    hostHeartbeatTimer = setTimeout(triggerHostDisconnect, 50000);
}

// =========================================================
// Temporary Bot Modal & Accessibility Management
// =========================================================
function setBackgroundInertForBotModal(isInert) {
    const bgIds = ['btn-home-top', 'game-screen', 'result-modal', 'waiting-room-screen', 'setup-screen', 'manual-screen', 'name-entry-screen'];
    bgIds.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        if (isInert) {
            el.setAttribute('aria-hidden', 'true');
            el.setAttribute('inert', '');
        } else {
            el.removeAttribute('aria-hidden');
            el.removeAttribute('inert');
        }
    });
}

function updateTemporaryBotModalCountdownText(seconds) {
    const cdEl = document.getElementById('temporary-bot-countdown');
    if (cdEl) {
        cdEl.innerText = `กรุณาเลือกภายใน ${seconds} วินาที มิฉะนั้นระบบจะถือว่าคุณเลือก "ไม่ใช่"`;
    }
}

function openTemporaryBotModal(initialSeconds = 30) {
    const modal = document.getElementById('temporary-bot-modal');
    if (!modal) return;

    hasSubmittedBotModalResponse = false;
    isLocalTempBotModalOpen = true;

    const btnYes = document.getElementById('btn-bot-return-yes');
    const btnNo = document.getElementById('btn-bot-return-no');
    if (btnYes) btnYes.disabled = false;
    if (btnNo) btnNo.disabled = false;

    updateTemporaryBotModalCountdownText(initialSeconds);
    setBackgroundInertForBotModal(true);
    modal.style.display = 'flex';

    if (localBotModalTimer) clearInterval(localBotModalTimer);
    let localSec = initialSeconds;
    localBotModalTimer = setInterval(() => {
        localSec--;
        if (localSec >= 0) {
            updateTemporaryBotModalCountdownText(localSec);
        }
        if (localSec <= 0) {
            clearInterval(localBotModalTimer);
            localBotModalTimer = null;
        }
    }, 1000);

    setTimeout(() => {
        const heading = document.getElementById('temporary-bot-heading');
        if (heading) heading.focus();
    }, 100);
}

function closeTemporaryBotModal() {
    const modal = document.getElementById('temporary-bot-modal');
    if (localBotModalTimer) {
        clearInterval(localBotModalTimer);
        localBotModalTimer = null;
    }
    isLocalTempBotModalOpen = false;
    if (modal) {
        modal.style.display = 'none';
    }
    setBackgroundInertForBotModal(false);
}

function setupTemporaryBotModalEvents() {
    const btnYes = document.getElementById('btn-bot-return-yes');
    const btnNo = document.getElementById('btn-bot-return-no');

    if (btnYes) {
        btnYes.onclick = () => handleBotModalChoice(true);
    }
    if (btnNo) {
        btnNo.onclick = () => handleBotModalChoice(false);
    }
}

function handleBotModalChoice(wantsToReturn) {
    if (hasSubmittedBotModalResponse) return;
    hasSubmittedBotModalResponse = true;

    const btnYes = document.getElementById('btn-bot-return-yes');
    const btnNo = document.getElementById('btn-bot-return-no');
    if (btnYes) btnYes.disabled = true;
    if (btnNo) btnNo.disabled = true;

    closeTemporaryBotModal();

    if (wantsToReturn) {
        if (isHost) {
            handlePlayerBotReturnYes(myPeerId);
        } else if (hostConnection && hostConnection.open) {
            hostConnection.send({ type: 'BOT_RETURN_YES', playerId: myPeerId });
        }
    } else {
        if (isHost) {
            handleHostRoomExit();
        } else {
            if (hostConnection && hostConnection.open) {
                try {
                    hostConnection.send({ type: 'CLIENT_LEAVE_ROOM', playerId: myPeerId });
                } catch (e) {}
            }
            setTimeout(() => {
                try { if (hostConnection) hostConnection.close(); } catch (e) {}
                try { if (peer) peer.destroy(); } catch (e) {}
                window.location.href = '../index.html';
            }, 150);
        }
    }
}

function clearBotResponseTimer(playerId) {
    if (botResponseTimers[playerId]) {
        clearInterval(botResponseTimers[playerId].interval);
        delete botResponseTimers[playerId];
    }
}

function startHostBotResponseTimer(playerId) {
    clearBotResponseTimer(playerId);
    let stateObj = {
        timeLeft: 30,
        responded: false,
        interval: null
    };
    botResponseTimers[playerId] = stateObj;

    if (playerId === myPeerId) {
        openTemporaryBotModal(30);
    } else {
        let conn = guestConnections.find(c => c.peer === playerId);
        if (conn) {
            conn.send({ type: 'SHOW_TEMP_BOT_MODAL', timeLeft: 30 });
        }
    }

    stateObj.interval = setInterval(() => {
        if (stateObj.responded) {
            clearBotResponseTimer(playerId);
            return;
        }
        stateObj.timeLeft--;

        if (playerId === myPeerId) {
            if (stateObj.timeLeft >= 0) updateTemporaryBotModalCountdownText(stateObj.timeLeft);
        } else {
            let conn = guestConnections.find(c => c.peer === playerId);
            if (conn && stateObj.timeLeft >= 0) {
                conn.send({ type: 'TEMP_BOT_MODAL_COUNTDOWN', timeLeft: stateObj.timeLeft });
            }
        }

        if (stateObj.timeLeft <= 0) {
            stateObj.responded = true;
            clearBotResponseTimer(playerId);
            handleBotModalTimeoutOnHost(playerId);
        }
    }, 1000);
}

function handleBotModalTimeoutOnHost(playerId) {
    if (playerId === myPeerId) {
        closeTemporaryBotModal();
        handleHostRoomExit();
    } else {
        let conn = guestConnections.find(c => c.peer === playerId);
        if (conn) {
            try {
                conn.send({ type: 'FORCE_LEAVE_TO_INDEX' });
            } catch (e) {}
        }
        handleClientDeparture(playerId, 'timeout');
    }
}

function handlePlayerBotReturnYes(playerId) {
    if (!isHost) return;
    let timerObj = botResponseTimers[playerId];
    if (timerObj) {
        if (timerObj.responded) return;
        timerObj.responded = true;
        clearBotResponseTimer(playerId);
    }

    let p = gameState.players.find(x => x.id === playerId);
    if (!p || p.permanentBot || !p.temporaryBot) return;
    if (p.returnRequested) return;

    p.returnRequested = true;

    if (playerId === myPeerId) {
        closeTemporaryBotModal();
    } else {
        let conn = guestConnections.find(c => c.peer === playerId);
        if (conn) {
            conn.send({ type: 'CLOSE_TEMP_BOT_MODAL' });
        }
    }

    syncStateToAll();
}

// =========================================================
// Unified Client Departure & Host Room Exit Handlers
// =========================================================
function handleClientDeparture(playerId, reason = 'disconnect') {
    if (!isHost || !playerId || playerId === myPeerId) return;
    if (processedDisconnects.has(playerId)) return;
    processedDisconnects.add(playerId);

    clearBotResponseTimer(playerId);
    pendingPeerSoundPayloads.delete(playerId);

    // 1. ปิดและลบ Connection ออกจาก guestConnections อย่างปลอดภัย
    let connIdx = guestConnections.findIndex(c => c.peer === playerId);
    if (connIdx !== -1) {
        let conn = guestConnections[connIdx];
        guestConnections.splice(connIdx, 1);
        try {
            conn.close();
        } catch (e) {}
    }

    // 2. กรณีอยู่ในห้องรอ (WAITING) ก่อนเริ่มเกม
    if (gameState.status === 'WAITING') {
        let rpIdx = roomPlayers.findIndex(p => p.id === playerId);
        let removedName = resolveName(playerId);
        if (rpIdx !== -1) {
            roomPlayers.splice(rpIdx, 1);
        }
        let joIdx = window.joinOrder.indexOf(playerId);
        if (joIdx !== -1) {
            window.joinOrder.splice(joIdx, 1);
        }
        broadcastLobby();
        broadcastAnnounce(`${removedName} ออกจากห้องรอผู้เล่น`);
        return;
    }

    // 3. กรณีอยู่ในเกม เปลี่ยนผู้เล่นคนนั้นเป็นบอทถาวร โดยรักษา Player ID, ชื่อ, ไพ่, คะแนน และลำดับเดิม
    let p = gameState.players.find(x => x.id === playerId);
    if (!p) return;

    p.isBot = true;
    p.originalIsBot = false;
    p.temporaryBot = false;
    p.returnRequested = false;
    p.permanentBot = true;

    broadcastAnnounce(`[PID:${p.id}] ออกจากห้อง ระบบเปลี่ยนเป็นบอทเล่นแทนจนจบเกม`);

    // หากอยู่ในช่วง PRE_GAME ให้ตรวจสอบว่าผู้เล่นจริงที่เหลือโหวตข้ามครบหรือยัง
    if (gameState.status === 'PRE_GAME') {
        syncStateToAll();
        let realPlayers = gameState.players.filter(x => !x.isOut && !x.isBot);
        if (realPlayers.length === 0 || (gameState.skipPreVotes && realPlayers.every(rp => gameState.skipPreVotes.includes(rp.id)))) {
            clearInterval(turnTimerInterval);
            triggerSound('60');
            gameState.status = 'PLAYING';
            let firstPlayer = gameState.players[gameState.turnIndex] || gameState.players[0];
            broadcastAnnounce(`📢 ผู้เล่นทุกคนพร้อมแล้ว! ขอเชิญ [PID:${firstPlayer.id}] จั่วไพ่ใบแรกและทิ้งลงกองเพื่อเริ่มเกมได้เลยครับ`);
            syncStateToAll();
            startTurnTimer();
            checkBotTurn();
        }
        return;
    }

    syncStateToAll();

    // หากเป็นตาของผู้เล่นที่เพิ่งออกในสถานะ PLAYING ให้บอทเข้าควบคุมต่อทันทีโดยเกมไม่ค้าง
    if (gameState.status === 'PLAYING') {
        let currentTurnPlayer = gameState.players[gameState.turnIndex];
        if (currentTurnPlayer && currentTurnPlayer.id === p.id) {
            clearInterval(turnTimerInterval);
            checkBotTurn();
        }
    }
}

function handleHostRoomExit() {
    if (!isHost || isClosingHostRoom) return;
    isClosingHostRoom = true;

    clearInterval(turnTimerInterval);
    clearInterval(heartbeatInterval);
    if (localBotModalTimer) {
        clearInterval(localBotModalTimer);
        localBotModalTimer = null;
    }
    Object.keys(botResponseTimers).forEach(pid => clearBotResponseTimer(pid));

    if (roomId) {
        try {
            remove(ref(db, GAME_DB_ID + '/rooms/' + roomId));
        } catch (e) {}
    }

    guestConnections.forEach(c => {
        try {
            c.send({ type: 'HOST_CLOSED_ROOM' });
        } catch (e) {}
    });

    setTimeout(() => {
        guestConnections.forEach(c => {
            try { c.close(); } catch (e) {}
        });
        guestConnections = [];
        try {
            if (peer) peer.destroy();
        } catch (e) {}
        window.location.href = '../index.html';
    }, 200);
}

let currentTopCardText = "ยังไม่มีไพ่บนกอง";

// --- CHEAT MODE BUFFERS ---
let cheatBuffer = "";
let cheatTimeout = null;

// Shortcut Keys & Cheat Code Event Listener
document.addEventListener('keydown', (e) => {
    if (isLocalTempBotModalOpen) return;

    // 1. Cheat Mode Logic
    if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'TEXTAREA') {
        if (e.key.length === 1 && !e.altKey && !e.ctrlKey && !e.metaKey) {
            cheatBuffer += e.key;
            clearTimeout(cheatTimeout);
            cheatTimeout = setTimeout(() => { cheatBuffer = ""; }, 2000);

            const cmd = cheatBuffer.toLowerCase();
            if (cmd.includes("ตอง")) { setCheatHand("ตอง"); cheatBuffer = ""; }
            else if (cmd.includes("ดอก") || cmd.includes("สี")) { setCheatHand("ดอก"); cheatBuffer = ""; }
            else if (cmd.includes("เรียง")) { setCheatHand("เรียง"); cheatBuffer = ""; }
            else if (cmd.includes("50")) { setCheatHand("50"); cheatBuffer = ""; }
            else if (cmd.includes("3a")) { setCheatHand("3a"); cheatBuffer = ""; }
        }
    }

    // 2. Shortcut Keys Logic
    if (e.altKey) {
        const key = e.key.toLowerCase();
        if (key === 'e') {
            e.preventDefault();
            let btn = document.getElementById('btn-end-turn');
            if(btn && !btn.disabled && btn.style.display !== 'none') btn.click();
        } else if (key === 'p') {
            e.preventDefault();
            let btn = document.getElementById('btn-draw');
            if(btn && !btn.disabled && btn.style.display !== 'none') btn.click();
        } else if (key === 'k') {
            e.preventDefault();
            let btn = document.getElementById('btn-kang');
            if(btn && !btn.disabled && btn.style.display !== 'none') btn.click();
        } else if (key === 'l') {
            e.preventDefault();
            let btn = document.getElementById('btn-kang25');
            if(btn && !btn.disabled && btn.style.display !== 'none') btn.click();
        } else if (key === 'c') {
            e.preventDefault();
            announce(currentTopCardText);
        }
    }
});

function updateHomeBtnVisibility() {
    const homeBtn = document.getElementById('btn-home-top');
    const nameScreen = document.getElementById('name-entry-screen');
    const nameDisp = nameScreen ? window.getComputedStyle(nameScreen).display : 'none';
    const setupDisp = window.getComputedStyle(document.getElementById('setup-screen')).display;
    if (setupDisp !== 'none' || nameDisp !== 'none') homeBtn.style.display = 'block';
    else homeBtn.style.display = 'none';
}

const bubblesBg = document.getElementById('bg-bubbles');
for(let i=0; i<15; i++) {
    let b = document.createElement('div'); b.className = 'bubble';
    b.style.width = Math.random() * 20 + 10 + 'px'; b.style.height = b.style.width;
    b.style.left = Math.random() * 100 + '%';
    b.style.animationDuration = Math.random() * 4 + 4 + 's';
    bubblesBg.appendChild(b);
}

function resolveName(id) {
    let p = globalPlayersMap[id];
    if (p && p.name && String(p.name).trim() !== '') {
        return String(p.name).trim();
    }
    if (gameState && gameState.players) {
        let gp = gameState.players.find(x => x.id === id);
        if (gp && gp.name && String(gp.name).trim() !== '') {
            return String(gp.name).trim();
        }
    }
    if (roomPlayers) {
        let rp = roomPlayers.find(x => x.id === id);
        if (rp && rp.name && String(rp.name).trim() !== '') {
            return String(rp.name).trim();
        }
    }
    if (typeof id === 'string' && id.startsWith('bot_')) {
        let bIdx = parseInt(id.split('_')[1]);
        if (!isNaN(bIdx) && botNames[bIdx]) return botNames[bIdx];
        return BOT_NAMES_POOL[bIdx % BOT_NAMES_POOL.length] || "บอท";
    }
    if (id === myPeerId && myPlayerName) return myPlayerName;
    return "ผู้เล่น";
}

// =========================================================
// UI Transitions & Announcements
// =========================================================
window.showTransition = (msg, duration) => {
    const overlay = document.getElementById('transition-overlay');
    const text = document.getElementById('transition-text');
    text.innerText = msg;
    overlay.style.display = 'flex';
    if (duration) setTimeout(() => { overlay.style.display = 'none'; }, duration);
};

window.hideTransition = () => { document.getElementById('transition-overlay').style.display = 'none'; };

function broadcastTransition(msg, duration) {
    if(myPeerId) window.showTransition(msg, duration);
    guestConnections.forEach(c => c.send({ type: 'SHOW_TRANSITION', msg, duration }));
}

function announce(rawMsg, showOnBanner = true) {
    let formattedMsg = rawMsg.replace(/\[PID:(.*?)\]/g, (match, id) => resolveName(id));
    const srOnly = document.getElementById('sr-only-announcer');
    if (srOnly) {
        srOnly.innerText = ''; 
        setTimeout(() => { srOnly.innerText = formattedMsg; }, 50);
    }
    if(showOnBanner && document.getElementById('game-screen').style.display === 'flex') {
        const visAnnouncer = document.getElementById('visible-game-announcer');
        if(visAnnouncer) {
            visAnnouncer.innerText = '';
            setTimeout(() => { visAnnouncer.innerText = formattedMsg; }, 50);
        }
    }
    if (announcerTimeout) clearTimeout(announcerTimeout);
    announcerTimeout = setTimeout(() => {
        const visAnnouncer = document.getElementById('visible-game-announcer');
        if (visAnnouncer) visAnnouncer.innerText = '';
        const srOnly = document.getElementById('sr-only-announcer');
        if (srOnly) srOnly.innerText = '';
    }, 5000);
}

function broadcastAnnounce(rawMsg) {
    let data = { type: 'ANNOUNCE', msg: rawMsg };
    if (myPeerId) announce(rawMsg); 
    guestConnections.forEach(c => c.send(data));
}

// UI Buttons (Manual, etc.)
window.openManual = () => {
    document.getElementById('setup-screen').style.display = 'none';
    document.getElementById('manual-screen').style.display = 'flex';
    updateHomeBtnVisibility();
    setTimeout(() => { document.getElementById('manual-title').focus(); }, 100);
    announce("เปิดหน้าคู่มือการเล่นแล้ว", false);
};

window.closeManual = () => {
    document.getElementById('manual-screen').style.display = 'none';
    document.getElementById('setup-screen').style.display = 'flex';
    updateHomeBtnVisibility();
    document.getElementById('btn-open-manual').focus();
    announce("ปิดหน้าคู่มือการเล่น กลับสู่หน้าหลัก", false);
};

const SUITS = ['♠', '♥', '♦', '♣'];
const RANKS = ['A','2','3','4','5','6','7','8','9','10','J','Q','K'];
function getCardValue(rank) {
    if(rank === 'A') return 1;
    if(['J','Q','K'].includes(rank)) return 10;
    return parseInt(rank);
}
function buildDeck() {
    let d = [];
    SUITS.forEach(s => RANKS.forEach(r => d.push({suit: s, rank: r, val: getCardValue(r)})));
    return d.sort(() => Math.random() - 0.5);
}

// =========================================================
// Player Name Entry Logic
// =========================================================
function setupNameEntryEvents() {
    const nameInput = document.getElementById('player-name-input');
    const confirmBtn = document.getElementById('btn-confirm-name');
    const nameScreen = document.getElementById('name-entry-screen');
    const setupScreen = document.getElementById('setup-screen');

    if (!nameInput || !confirmBtn || !nameScreen) return;

    const handleConfirmName = (e) => {
        if (e) e.preventDefault();
        if (isNameConfirmed) return;

        const trimmedName = nameInput.value.trim();
        if (!trimmedName) {
            announce("กรุณากรอกชื่อผู้เล่นก่อนเข้าสู่เกม", false);
            nameInput.focus();
            return;
        }

        isNameConfirmed = true;
        myPlayerName = trimmedName;
        confirmBtn.disabled = true;
        nameInput.disabled = true;

        playSound('select');

        nameScreen.style.display = 'none';
        if (setupScreen) {
            setupScreen.style.display = 'flex';
        }
        updateHomeBtnVisibility();
        announce(`ยินดีต้อนรับ ${myPlayerName} เข้าสู่หน้าเลือกห้องเล่นแคง`, false);

        setTimeout(() => {
            const createRoomBtn = document.getElementById('btn-create-room');
            if (createRoomBtn) createRoomBtn.focus();
        }, 100);
    };

    confirmBtn.addEventListener('click', handleConfirmName);
    nameInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            handleConfirmName(e);
        }
    });
}

// =========================================================
// Network & Lobby Logic
// =========================================================
window.onload = () => {
    myPeerId = "cang-client-" + Math.random().toString(36).substring(2, 8);
    peer = new Peer(myPeerId);
    setupNameEntryEvents();
    setupTemporaryBotModalEvents();
    updateHomeBtnVisibility();

    const roomsRef = ref(db, GAME_DB_ID + '/rooms');
    onValue(roomsRef, (snapshot) => {
        const listContainer = document.getElementById('room-list-container');
        listContainer.innerHTML = '';
        if (snapshot.exists()) {
            let hasRoom = false;
            for (let rId in snapshot.val()) {
                let info = snapshot.val()[rId];
                if (info.state === 'WAITING' && info.playerCount < 6) {
                    hasRoom = true;
                    let btn = document.createElement('button');
                    btn.className = 'room-item-btn';
                    btn.setAttribute('aria-label', `เข้าร่วมห้อง ${rId} มีผู้เล่น ${info.playerCount} จาก 6 คน`);
                    btn.innerHTML = `ห้อง ${rId} (${info.playerCount}/6 คน) <span aria-hidden="true">เข้า ➔</span>`;
                    btn.onclick = () => window.joinRoomDirect(rId);
                    listContainer.appendChild(btn);
                }
            }
            if(!hasRoom) listContainer.innerHTML = '<div class="no-room-text" role="status">ไม่มีห้องว่าง (กดสร้างห้องได้เลย)</div>';
        } else {
            listContainer.innerHTML = '<div class="no-room-text" role="status">ไม่มีห้องว่าง (กดสร้างห้องได้เลย)</div>';
        }
    });
};

window.createNewRoom = () => {
    if (!myPlayerName || !myPlayerName.trim()) return;
    isHost = true; announce("กำลังสร้างห้อง โปรดรอสักครู่", false);
    const seqRef = ref(db, GAME_DB_ID + '/room_seq');
    runTransaction(seqRef, (currentData) => {
        let nextNum = (currentData || 0) + 1;
        if(nextNum > 99999) nextNum = 1; return nextNum;
    }).then((result) => {
        if (result.committed) {
            roomId = 'Cang' + String(result.snapshot.val()).padStart(5, '0');
            if (peer) peer.destroy();
            peer = new Peer("cang-host-" + roomId);
            peer.on('open', (id) => {
                myPeerId = id; selectedColorId = 'red';
                window.joinOrder = [myPeerId]; 
                roomPlayers = [{ id: myPeerId, name: myPlayerName, color: selectedColorId }];
                botColors = [];
                botNames = [];
                currentBotCount = 0;
                processedDisconnects.clear();
                const myRoomRef = ref(db, GAME_DB_ID + '/rooms/' + roomId);
                set(myRoomRef, { roomId: roomId, playerCount: 1, state: 'WAITING' });
                onDisconnect(myRoomRef).remove();
                document.getElementById('setup-screen').style.display = 'none';
                document.getElementById('waiting-room-screen').style.display = 'flex';
                updateHomeBtnVisibility();
                announce(`สร้างห้องสำเร็จ รหัสห้องคือ ${roomId}`, false);
                playSound('1');
                updateLobbyUI(); listenHostConnections();

                heartbeatInterval = setInterval(() => { guestConnections.forEach(c => c.send({ type: 'HEARTBEAT' })); }, 15000);
            });
        }
    });
};

window.joinRoomDirect = (targetRoomId) => {
    if (!myPlayerName || !myPlayerName.trim()) return;
    isHost = false; roomId = targetRoomId;
    announce(`กำลังเข้าร่วมห้อง ${roomId}`, false);
    document.getElementById('setup-screen').style.display = 'none';
    document.getElementById('waiting-room-screen').style.display = 'flex';
    document.getElementById('host-bot-section').style.display = 'none';
    updateHomeBtnVisibility();
    
    hostConnection = peer.connect("cang-host-" + roomId);
    hostConnection.on('open', () => { 
        hostConnection.send({ type: 'JOIN_ROOM', name: myPlayerName }); 
        announce(`เข้าร่วมห้องสำเร็จ รอแจกไพ่`, false); 
        resetHeartbeat();
    });
    hostConnection.on('data', (data) => {
        resetHeartbeat();
        if (data.type === 'HEARTBEAT') return;
        if (data.type === 'HOST_CLOSED_ROOM') {
            triggerHostDisconnect();
            return;
        }
        if (data.type === 'FORCE_LEAVE_TO_INDEX') {
            closeTemporaryBotModal();
            try { if (hostConnection) hostConnection.close(); } catch (e) {}
            try { if (peer) peer.destroy(); } catch (e) {}
            window.location.href = '../index.html';
            return;
        }
        if (data.type === 'SHOW_TEMP_BOT_MODAL') {
            openTemporaryBotModal(data.timeLeft || 30);
            return;
        }
        if (data.type === 'TEMP_BOT_MODAL_COUNTDOWN') {
            if (localBotModalTimer) {
                clearInterval(localBotModalTimer);
                localBotModalTimer = null;
            }
            updateTemporaryBotModalCountdownText(data.timeLeft);
            return;
        }
        if (data.type === 'CLOSE_TEMP_BOT_MODAL') {
            closeTemporaryBotModal();
            return;
        }
        if (data.type === 'LOBBY_UPDATE') { 
            roomPlayers = data.roomPlayers || [];
            currentBotCount = data.botCount || 0;
            botColors = data.botColors || [];
            botNames = data.botNames || [];
            if (data.joinOrder) window.joinOrder = data.joinOrder; 
            updateLobbyUI(); 
        }
        if (data.type === 'START_COUNTDOWN') { 
            if (!window.isStartingRound) { window.isStartingRound = true; doCountdownAndStart(); }
        }
        if (data.type === 'SYNC_STATE') renderClientGame(data.publicState, data.privateState);
        if (data.type === 'ANIM_SC_FLOAT') showFloatingScore(data.targetId, data.amt);
        if (data.type === 'SHOW_RESULT') showResultModal(data.title, data.detail, data.isEnd, data.winners, data.losers);
        if (data.type === 'ANNOUNCE') announce(data.msg);
        if (data.type === 'TIMER_UPDATE') window.updateTimerUI(data.timeLeft, data.maxTime);
        if (data.type === 'PLAY_SOUND') handleIncomingPlaySoundData(data);
        if (data.type === 'SHOW_TRANSITION') window.showTransition(data.msg, data.duration);
        if (data.type === 'HIDE_TRANSITION') window.hideTransition();
        if (data.type === 'KANG_ANIMATION') window.playKangAnimation(data.callerId, data.winType);
    });
    hostConnection.on('close', () => {
        triggerHostDisconnect();
    });
    hostConnection.on('error', () => {
        triggerHostDisconnect();
    });
};

function listenHostConnections() {
    peer.on('connection', (conn) => {
        conn.on('data', (data) => {
            if (data.type === 'JOIN_ROOM') {
                if (roomPlayers.length + currentBotCount >= 6) return;
                let usedColors = roomPlayers.map(p => p.color).concat(botColors);
                let avail = COLORS.filter(c => !usedColors.includes(c.id));
                let randomColor = avail.length > 0 ? avail[Math.floor(Math.random() * avail.length)].id : 'red';
                let joinedName = (data.name && String(data.name).trim() !== '') ? String(data.name).trim() : 'ผู้เล่น';
                
                roomPlayers.push({ id: conn.peer, name: joinedName, color: randomColor });
                window.joinOrder.push(conn.peer); 
                guestConnections.push(conn); 
                broadcastLobby();
                broadcastAnnounce(`${joinedName} เข้าร่วมห้องใหม่ รวม ${roomPlayers.length + currentBotCount} คน`);
                triggerSound('select');
            }
            if (data.type === 'PLAYER_ACTION') processPlayerAction(data.playerId, data.action, data.cardIndex);
            if (data.type === 'BOT_RETURN_YES') handlePlayerBotReturnYes(data.playerId || conn.peer);
            if (data.type === 'CLIENT_LEAVE_ROOM') handleClientDeparture(data.playerId || conn.peer, 'voluntary');
        });
        conn.on('close', () => {
            handleClientDeparture(conn.peer, 'disconnect');
        });
        conn.on('error', () => {
            handleClientDeparture(conn.peer, 'error');
        });
    });
}

function broadcastLobby() {
    let data = { type: 'LOBBY_UPDATE', roomPlayers, botCount: currentBotCount, botColors, botNames, joinOrder: window.joinOrder };
    guestConnections.forEach(c => c.send(data)); updateLobbyUI();
    if(isHost) set(ref(db, GAME_DB_ID + '/rooms/' + roomId + '/playerCount'), roomPlayers.length + currentBotCount);
}

window.adjustBot = (delta) => {
    if(!isHost) return;
    let currentTotal = roomPlayers.length + currentBotCount;
    if (delta > 0 && currentTotal < 6) {
        let used = roomPlayers.map(p => p.color).concat(botColors);
        let avail = COLORS.filter(c => !used.includes(c.id));
        let randomColorId = avail.length > 0 ? avail[Math.floor(Math.random() * avail.length)].id : 'red';
        
        botColors.push(randomColorId); window.joinOrder.push('bot_' + currentBotCount); currentBotCount++;
        randomizeAllBotNames(currentBotCount);
    } else if (delta < 0 && currentBotCount > 0) {
        currentBotCount--; botColors.pop(); 
        for (let i = window.joinOrder.length - 1; i >= 0; i--) {
            if (window.joinOrder[i].startsWith('bot_')) { window.joinOrder.splice(i, 1); break; }
        }
        randomizeAllBotNames(currentBotCount);
    }
    broadcastLobby(); triggerSound('select');
};

function updateLobbyUI() {
    document.getElementById('display-room-id').innerText = roomId;
    document.getElementById('bot-count-text').innerText = `${currentBotCount} ตัว`;
    
    globalPlayersMap = {};
    roomPlayers.forEach(p => {
        let c = COLORS.find(x => x.id === p.color) || COLORS[0];
        let pName = (p.name && String(p.name).trim() !== '') ? String(p.name).trim() : (p.id === myPeerId && myPlayerName ? myPlayerName : 'ผู้เล่น');
        globalPlayersMap[p.id] = { isBot: false, name: pName, colorName: c ? c.name : 'แดง', colorHex: c ? c.hex : 'var(--red)' };
    });
    botColors.forEach((b, i) => {
        let c = COLORS.find(x => x.id === b) || COLORS[i % COLORS.length];
        let bName = botNames[i] || BOT_NAMES_POOL[i % BOT_NAMES_POOL.length];
        globalPlayersMap['bot_' + i] = { isBot: true, name: bName, colorName: c ? c.name : 'แดง', colorHex: c ? c.hex : 'var(--red)' };
    });

    const listEl = document.getElementById('waiting-players-list'); listEl.innerHTML = '';
    
    let currentOrder = window.joinOrder && window.joinOrder.length > 0 ? window.joinOrder : roomPlayers.map(p => p.id).concat(botColors.map((_, i) => 'bot_'+i));
    currentOrder.forEach(id => {
        if (id.startsWith('bot_')) {
            let li = document.createElement('li'); li.innerText = resolveName(id); listEl.appendChild(li);
        } else {
            let p = roomPlayers.find(x => x.id === id);
            if (p) { let li = document.createElement('li'); li.innerText = resolveName(p.id); listEl.appendChild(li); }
        }
    });

    if(isHost) {
        let currentTotal = roomPlayers.length + currentBotCount;
        let btnDec = document.getElementById('btn-bot-dec'); let btnInc = document.getElementById('btn-bot-inc');
        if (btnDec) btnDec.disabled = (currentBotCount === 0);
        if (btnInc) btnInc.disabled = (currentTotal >= 6);
        document.getElementById('btn-host-start').style.display = 'block';
        document.getElementById('btn-host-start').disabled = (roomPlayers.length + currentBotCount < 3);
    } else { document.getElementById('guest-waiting-text').style.display = 'block'; }
}

window.hostStartGame = () => {
    if (window.isStartingRound) return;
    window.isStartingRound = true;
    let startBtn = document.getElementById('btn-host-start');
    if (startBtn) { startBtn.disabled = true; startBtn.style.display = 'none'; }
    remove(ref(db, GAME_DB_ID + '/rooms/' + roomId)); 
    guestConnections.forEach(c => sendSoundMessageToConn(c, { type: 'START_COUNTDOWN' }));
    doCountdownAndStart();
};

function doCountdownAndStart() {
    document.getElementById('waiting-room-screen').style.display = 'none';
    document.getElementById('result-modal').style.display = 'none';
    updateHomeBtnVisibility();

    const countdownScreen = document.getElementById('countdown-screen');
    countdownScreen.style.display = 'flex';
    const cdNum = document.getElementById('cd-number');
    const srCD = document.getElementById('sr-countdown');
    
    playSound('start'); 
    let count = 3; cdNum.innerText = count; srCD.innerText = count;
    
    let cdInterval = setInterval(() => {
        count--;
        if (count > 0) { cdNum.innerText = count; srCD.innerText = count; } 
        else {
            clearInterval(cdInterval); countdownScreen.style.display = 'none'; document.getElementById('game-screen').style.display = 'flex'; playBGM(); 
            if (isHost) {
                if (gameState.status === 'WAITING') {
                    gameState.players = [];
                    window.joinOrder.forEach(id => {
                        if (id.startsWith('bot_')) {
                            let botIdx = parseInt(id.split('_')[1]);
                            let bColor = botColors[botIdx] || COLORS[botIdx % COLORS.length].id;
                            let bName = botNames[botIdx] || BOT_NAMES_POOL[botIdx % BOT_NAMES_POOL.length];
                            let cInfo = COLORS.find(c => c.id === bColor) || COLORS[0];
                            gameState.players.push({
                                id: id,
                                name: bName,
                                colorInfo: cInfo,
                                isBot: true,
                                originalIsBot: true,
                                temporaryBot: false,
                                returnRequested: false,
                                permanentBot: false,
                                consecutiveTimeouts: 0,
                                nextTurnTimeLimit: 40,
                                points: 200,
                                hand: [],
                                isOut: false,
                                turnCount: 0,
                                hasFinishedFirstTurn: false,
                                hasFlowedThisTurn: false
                            });
                        } else {
                            let p = roomPlayers.find(x => x.id === id);
                            if (p) {
                                let cInfo = COLORS.find(c => c.id === p.color) || COLORS[0];
                                let pName = (p.name && String(p.name).trim() !== '') ? String(p.name).trim() : (p.id === myPeerId && myPlayerName ? myPlayerName : 'ผู้เล่น');
                                gameState.players.push({
                                    id: p.id,
                                    name: pName,
                                    colorInfo: cInfo,
                                    isBot: false,
                                    originalIsBot: false,
                                    temporaryBot: false,
                                    returnRequested: false,
                                    permanentBot: false,
                                    consecutiveTimeouts: 0,
                                    nextTurnTimeLimit: 40,
                                    points: 200,
                                    hand: [],
                                    isOut: false,
                                    turnCount: 0,
                                    hasFinishedFirstTurn: false,
                                    hasFlowedThisTurn: false
                                });
                            }
                        }
                    });
                }
                startNewRound();
            }
        }
    }, 1630);
}

// =========================================================
// Game Logic Core
// =========================================================
function restorePlayerIfReturnRequested(p) {
    if (!p) return false;
    if (p.temporaryBot && p.returnRequested && !p.permanentBot && !p.originalIsBot) {
        p.isBot = false;
        p.temporaryBot = false;
        p.returnRequested = false;
        p.consecutiveTimeouts = 0;
        p.nextTurnTimeLimit = 40;
        return true;
    }
    return false;
}

function startNewRound() {
    if (window.nextRoundHostId) {
        let idx = gameState.players.findIndex(p => p.id === window.nextRoundHostId);
        if (idx > 0) {
            let newHost = gameState.players.splice(idx, 1)[0];
            gameState.players.unshift(newHost);
        }
    }

    deck = buildDeck(); discardPile = [];
    gameState.status = 'PRE_GAME'; gameState.turnIndex = 0; gameState.topCardOwnerId = null; gameState.flowSourceId = null; gameState.skipPreVotes = []; 
    isProcessingTurnTimeout = false;
    isEndingTurn = false;
    activeBotSequencePlayerId = null;
    
    gameState.players.forEach((p) => {
        if(p.points > 0) { 
            p.hand = deck.splice(0, 5); p.isOut = false; p.hasDrawnTurn = false; p.hasDiscardedTurn = false; p.hasFlowedThisTurn = false; p.discardedRankThisTurn = null; p.turnCount = 0; p.hasFinishedFirstTurn = false;
        } else { p.isOut = true; } 
    });
    
    syncStateToAll(); startPreGameTimer(); broadcastAnnounce("เปิดรอบใหม่! มีใครจะแคง 25 ไหม?");
}

function startPreGameTimer() {
    clearInterval(turnTimerInterval);
    currentTurnMaxTime = 40;
    let timeLeft = 40; broadcastTimer(timeLeft, 40);
    turnTimerInterval = setInterval(() => {
        timeLeft--; broadcastTimer(timeLeft, 40);
        if(timeLeft <= 0) { 
            clearInterval(turnTimerInterval); 
            if (gameState.status === 'PRE_GAME') {
                triggerSound('60'); gameState.status = 'PLAYING';
                let firstPlayer = gameState.players[gameState.turnIndex] || gameState.players[0];
                restorePlayerIfReturnRequested(firstPlayer);
                broadcastAnnounce(`📢 หมดเวลาดูไพ่ 40 วินาที ไม่มีใครกดแคง! ขอเชิญ [PID:${firstPlayer.id}] จั่วไพ่ใบแรกและทิ้งลงกองเพื่อเริ่มเกมได้เลยครับ`);
                syncStateToAll(); startTurnTimer(); checkBotTurn();
            }
        }
    }, 1000);
}

function startTurnTimer() {
    clearInterval(turnTimerInterval);
    let cp = gameState.players[gameState.turnIndex];
    let maxTime = (cp && !cp.isBot && cp.nextTurnTimeLimit === 20) ? 20 : 40;
    currentTurnMaxTime = maxTime;
    let timeLeft = maxTime;
    broadcastTimer(timeLeft, maxTime);
    turnTimerInterval = setInterval(() => {
        timeLeft--; broadcastTimer(timeLeft, maxTime);
        if(timeLeft <= 0) { clearInterval(turnTimerInterval); handleTurnTimeout(); }
    }, 1000);
}

function broadcastTimer(t, maxT = currentTurnMaxTime) {
    if(myPeerId) window.updateTimerUI(t, maxT);
    guestConnections.forEach(c => c.send({ type: 'TIMER_UPDATE', timeLeft: t, maxTime: maxT }));
}

window.updateTimerUI = (timeLeft, maxTime = 40) => {
    const timerBar = document.getElementById('timer-bar-fill');
    if(timerBar) {
        let validMax = (maxTime && maxTime > 0) ? maxTime : 40;
        let pct = Math.max(0, (timeLeft / validMax) * 100);
        timerBar.style.width = pct + '%'; timerBar.style.background = timeLeft <= 10 ? 'var(--red)' : 'var(--orange)';
    }
};

function handleTurnTimeout() {
    if (isProcessingTurnTimeout) return;
    let cp = gameState.players[gameState.turnIndex];
    if (!cp || cp.isBot || gameState.status !== 'PLAYING') return;

    isProcessingTurnTimeout = true;
    try {
        cp.consecutiveTimeouts = (cp.consecutiveTimeouts || 0) + 1;

        if (cp.consecutiveTimeouts === 1) {
            // ระดับที่ 1: หมดเวลาครั้งที่ 1
            cp.nextTurnTimeLimit = 40;
            broadcastAnnounce(`⏰ [PID:${cp.id}] หมดเวลาเล่น (ครั้งที่ 1/3) ระบบเล่นแทนอัตโนมัติ`);
            if (discardPile.length === 0 && !cp.hasDiscardedTurn) {
                if (!cp.hasDrawnTurn) processPlayerAction(cp.id, 'DRAW', 0, true);
                if (cp.hand.length > 0) processPlayerAction(cp.id, 'DISCARD', 0, true);
            } else {
                if (!cp.hasDrawnTurn) processPlayerAction(cp.id, 'DRAW', 0, true);
                if (!cp.hasDiscardedTurn && cp.hand.length > 0) processPlayerAction(cp.id, 'DISCARD', 0, true);
            }
            if (gameState.status === 'PLAYING') processPlayerAction(cp.id, 'END_TURN', 0, true);
        } else if (cp.consecutiveTimeouts === 2) {
            // ระดับที่ 2: หมดเวลาครั้งที่ 2 ติดต่อกัน ลดเวลาตาถัดไปเหลือ 20 วินาที
            cp.nextTurnTimeLimit = 20;
            broadcastAnnounce(`⏰ [PID:${cp.id}] หมดเวลาเล่น 2 ครั้งติดต่อกัน! ตาถัดไปจะลดเวลาเหลือ 20 วินาที`);
            if (discardPile.length === 0 && !cp.hasDiscardedTurn) {
                if (!cp.hasDrawnTurn) processPlayerAction(cp.id, 'DRAW', 0, true);
                if (cp.hand.length > 0) processPlayerAction(cp.id, 'DISCARD', 0, true);
            } else {
                if (!cp.hasDrawnTurn) processPlayerAction(cp.id, 'DRAW', 0, true);
                if (!cp.hasDiscardedTurn && cp.hand.length > 0) processPlayerAction(cp.id, 'DISCARD', 0, true);
            }
            if (gameState.status === 'PLAYING') processPlayerAction(cp.id, 'END_TURN', 0, true);
        } else if (cp.consecutiveTimeouts >= 3) {
            // ระดับที่ 3: หมดเวลาครบ 3 ครั้ง เปลี่ยนเป็นบอทชั่วคราวทันทีและเปิด Modal เฉพาะเครื่องนั้น
            cp.isBot = true;
            cp.temporaryBot = true;
            cp.returnRequested = false;
            cp.nextTurnTimeLimit = 40;

            broadcastAnnounce(`🤖 [PID:${cp.id}] หมดเวลาครบ 3 ครั้ง ระบบเปลี่ยนเป็นบอทชั่วคราวเพื่อให้เกมดำเนินต่อไป`);
            syncStateToAll();
            startHostBotResponseTimer(cp.id);

            // ให้บอทควบคุมตัวละครนั้นต่อทันทีในตานี้โดยไม่หยุดเกม
            checkBotTurn();
        }
    } finally {
        isProcessingTurnTimeout = false;
    }
}

function syncStateToAll() {
    let activePlayers = gameState.players.filter(p => !p.isOut);
    if(activePlayers.length <= 1) return;

    let publicState = {
        status: gameState.status, turnId: gameState.players[gameState.turnIndex].id, topDiscard: discardPile[discardPile.length - 1],
        deckCount: deck.length, discardPileCount: discardPile.length, skipVotes: gameState.skipPreVotes || [],
        playersInfo: gameState.players.map(p => ({
            id: p.id,
            name: p.name,
            colorHex: p.colorInfo ? p.colorInfo.hex : 'var(--red)',
            colorName: p.colorInfo ? p.colorInfo.name : 'แดง',
            isBot: p.isBot,
            temporaryBot: !!p.temporaryBot,
            returnRequested: !!p.returnRequested,
            permanentBot: !!p.permanentBot,
            points: p.points,
            cardCount: p.hand.length,
            isOut: p.isOut
        }))
    };

    gameState.players.forEach(p => {
        // ซิงก์ข้อมูลให้ผู้เล่นจริงทุกคน รวมถึงผู้เล่นที่ถูกเปลี่ยนเป็นบอทชั่วคราว (ตราบใดที่ยังไม่ได้ออกจากห้องถาวร)
        const isHumanSlot = !p.originalIsBot && !String(p.id).startsWith('bot_');
        if (isHumanSlot && !p.permanentBot) {
            let privateState = {
                hand: p.hand,
                hasDrawn: p.hasDrawnTurn,
                hasDiscarded: p.hasDiscardedTurn,
                discardedRank: p.discardedRankThisTurn,
                points: p.points,
                turnCount: p.turnCount,
                hasFinishedFirstTurn: p.hasFinishedFirstTurn,
                hasFlowedThisTurn: p.hasFlowedThisTurn,
                temporaryBot: !!p.temporaryBot,
                returnRequested: !!p.returnRequested
            };
            if (p.id === myPeerId) renderClientGame(publicState, privateState);
            else {
                let conn = guestConnections.find(c => c.peer === p.id);
                if(conn) conn.send({ type: 'SYNC_STATE', publicState, privateState });
            }
        }
    });
}

function broadcastFloatSC(targetId, amt) {
    let data = { type: 'ANIM_SC_FLOAT', targetId, amt };
    if(myPeerId) window.showFloatingScore(targetId, amt);
    guestConnections.forEach(c => c.send(data));
}

// =========================================================
// [พื้นที่สำหรับพัฒนาต่อ #2] ระบบ Animation (เพิ่มมิติการแสดงผล Knock / Special Win)
// =========================================================
window.playKangAnimation = (callerId, winType) => {
    const animScreen = document.getElementById('kang-animation-screen');
    const text1 = document.getElementById('kang-anim-text1');
    const text2 = document.getElementById('kang-anim-text2');
    const avatar = document.getElementById('kang-anim-avatar');

    animScreen.style.display = 'flex';
    
    // --- [เพิ่มมิติการแสดงผล] Visual Trigger: Screen Flash ---
    animScreen.style.animation = 'none';
    void animScreen.offsetWidth; // trigger reflow
    animScreen.style.animation = 'flashScreen 0.4s ease-out';
    
    text1.style.opacity = '0';
    text2.style.opacity = '0';
    
    let pInfo = globalPlayersMap[callerId];
    let pColorName = pInfo ? pInfo.colorName : '';
    let pColorHex = pInfo ? pInfo.colorHex : 'white';
    
    avatar.style.color = pColorHex;
    avatar.className = 'kang-anim-active';
    
    // --- [เพิ่มมิติการแสดงผล] Visual Trigger: Dramatic Avatar Pop ---
    avatar.style.transform = 'scale(0.1) rotate(-15deg)';
    avatar.style.transition = 'transform 0.6s cubic-bezier(0.34, 1.56, 0.64, 1)';
    setTimeout(() => avatar.style.transform = 'scale(1) rotate(0deg)', 50);

    // --- [เพิ่มมิติการแสดงผล] Visual Trigger: Screen Shake ---
    if (winType === 'cang25' || winType === 'knock' || winType === 'special') {
        document.body.style.animation = 'shake 0.5s ease-in-out';
        setTimeout(() => document.body.style.animation = '', 500);
    }

    let soundKey = '';
    if (winType === 'cang25') soundKey = 'cang25';
    else if (winType === 'knock') soundKey = 'knock';
    else if (winType === 'special') soundKey = 'special';
    else soundKey = 'cang';

    playSound(soundKey);

    let callerName = resolveName(callerId);

    setTimeout(() => {
        if (winType === 'cang') {
            text1.innerText = `โอ้ ${callerName} เอ่ยคำว่า แคง`;
        } else if (winType === 'cang25') {
            text1.innerText = `โอ้ ${callerName} เอ่ยคำว่า แคง25`;
        } else if (winType === 'knock') {
            text1.innerText = "KNOCK!";
        } else if (winType === 'special') {
            text1.innerText = "SPECIAL WIN!";
        }
        text1.style.opacity = '1'; announce(text1.innerText, false);
    }, 1200);

    setTimeout(() => {
        if (winType === 'cang' || winType === 'cang25') {
            text2.innerText = `เปิดไพ่ทุกคนเพื่อวัดดวงกัน`;
        } else if (winType === 'knock') {
            text2.innerText = `น็อกแล้วโดย ${callerName}`;
        } else if (winType === 'special') {
            text2.innerText = `กติกาพิเศษโดย ${callerName}`;
        }
        text2.style.opacity = '1'; announce(text2.innerText, false);
    }, 3000);

    setTimeout(() => {
        animScreen.style.display = 'none'; avatar.className = '';
    }, 4500);
};

function broadcastKangAnimation(callerId, winType) {
    gameState.status = 'TRANSITION'; clearInterval(turnTimerInterval); syncStateToAll();
    let data = { type: 'KANG_ANIMATION', callerId, winType };
    if (myPeerId) window.playKangAnimation(callerId, winType);
    guestConnections.forEach(c => {
        if (c && c.peer !== myPeerId) {
            sendSoundMessageToConn(c, data);
        }
    });
}
// =========================================================

function nextTurn() {
    isEndingTurn = false;
    activeBotSequencePlayerId = null;

    let prevPlayer = gameState.players[gameState.turnIndex];
    prevPlayer.hasDrawnTurn = false; prevPlayer.hasDiscardedTurn = false; prevPlayer.hasFlowedThisTurn = false; prevPlayer.discardedRankThisTurn = null;

    do { gameState.turnIndex = (gameState.turnIndex + 1) % gameState.players.length;
    } while (gameState.players[gameState.turnIndex].isOut || gameState.players[gameState.turnIndex].hand.length === 0);
    
    let currentTurnPlayer = gameState.players[gameState.turnIndex];
    if (restorePlayerIfReturnRequested(currentTurnPlayer)) {
        broadcastAnnounce(`✅ [PID:${currentTurnPlayer.id}] กลับมาควบคุมการเล่นตามปกติแล้ว`);
    }

    triggerSound('turn'); syncStateToAll(); startTurnTimer(); checkBotTurn();
}

function checkBotTurn() {
    let cp = gameState.players[gameState.turnIndex];
    if (!cp || !cp.isBot || gameState.status !== 'PLAYING') return;
    if (activeBotSequencePlayerId === cp.id) return;
    activeBotSequencePlayerId = cp.id;
    setTimeout(() => playBotSequence(cp), 1000);
}

async function playBotSequence(cp) {
    try {
        await delay(2000);
        if (gameState.status !== 'PLAYING' || gameState.players[gameState.turnIndex]?.id !== cp.id || !cp.isBot) return;

        let handVal = cp.hand.reduce((sum, c) => sum + c.val, 0);
        if (handVal <= 5 && discardPile.length > 0 && !cp.hasDrawnTurn && !cp.hasDiscardedTurn && !cp.hasFlowedThisTurn) {
            processPlayerAction(cp.id, 'KANG', 0, true);
            return;
        }

        let topCard = discardPile[discardPile.length - 1];
        if (topCard && !cp.hasDrawnTurn && !cp.hasDiscardedTurn) {
            let flowIndex = cp.hand.findIndex(c => c.rank === topCard.rank);
            if (flowIndex !== -1) {
                processPlayerAction(cp.id, 'FLOW', flowIndex, true);
                await delay(2500);
                while (gameState.status === 'PLAYING' && gameState.players[gameState.turnIndex]?.id === cp.id && cp.isBot) {
                    let matchIdx = cp.hand.findIndex(c => c.rank === topCard.rank);
                    if (matchIdx !== -1) {
                        processPlayerAction(cp.id, 'FLOW', matchIdx, true);
                        await delay(1500);
                    } else break;
                }
                await delay(2000);
                if (gameState.status === 'PLAYING' && gameState.players[gameState.turnIndex]?.id === cp.id && cp.isBot) {
                    processPlayerAction(cp.id, 'END_TURN', 0, true);
                }
                return;
            }
        }

        if (!cp.hasDrawnTurn && !cp.hasFlowedThisTurn) {
            processPlayerAction(cp.id, 'DRAW', 0, true);
            await delay(2000);
        }
        if (gameState.status !== 'PLAYING' || gameState.players[gameState.turnIndex]?.id !== cp.id || !cp.isBot) return;

        if (!cp.hasDiscardedTurn && !cp.hasFlowedThisTurn && cp.hand.length > 0) {
            let rankCounts = {};
            cp.hand.forEach(c => rankCounts[c.rank] = (rankCounts[c.rank] || 0) + 1);
            let bestRank = null, maxCount = 0, maxVal = -1;
            for (let r in rankCounts) {
                let count = rankCounts[r];
                let val = getCardValue(r);
                if (count > maxCount || (count === maxCount && val > maxVal)) {
                    maxCount = count;
                    maxVal = val;
                    bestRank = r;
                }
            }
            while (gameState.status === 'PLAYING' && gameState.players[gameState.turnIndex]?.id === cp.id && cp.isBot) {
                let matchIdx = cp.hand.findIndex(c => c.rank === bestRank);
                if (matchIdx !== -1) {
                    processPlayerAction(cp.id, 'DISCARD', matchIdx, true);
                    await delay(1500);
                } else break;
            }
        }

        await delay(2000);
        if (gameState.status === 'PLAYING' && gameState.players[gameState.turnIndex]?.id === cp.id && cp.isBot) {
            processPlayerAction(cp.id, 'END_TURN', 0, true);
        }
    } finally {
        if (activeBotSequencePlayerId === cp.id) {
            activeBotSequencePlayerId = null;
        }
    }
}

function resolveSpecialWin(callerId, winType) {
    // ใช้งาน Special Win Animation แทรกก่อนทำการสรุปผลและคำนวณแต้ม
    broadcastKangAnimation(callerId, 'special');
    
    setTimeout(() => {
        window.nextRoundHostId = callerId;
        gameState.status = 'TRANSITION'; 
        clearInterval(turnTimerInterval);
        
        let active = gameState.players.filter(p => !p.isOut);
        let beforePoints = {};
        active.forEach(p => beforePoints[p.id] = p.points);
        
        let caller = active.find(p => p.id === callerId);
        let others = active.filter(p => p.id !== callerId);
        
        let aceCount = caller.hand.filter(c => c.rank === 'A').length;
        let payPerPerson = 0;
        
        if (winType === '3A') { payPerPerson = 40; } else { payPerPerson = 20 + (aceCount * 10); }
        
        others.forEach(p => { p.points -= payPerPerson; if(p.points < 0) p.points = 0; caller.points += payPerPerson; });
        
        let title = `[PID:${caller.id}] ชนะพิเศษกติกา: ${winType}!`;
        let winners = [caller.id]; let losers = [];
        let cardDetailLines = [`[PID:${caller.id}] เปิดไพ่ชนะกติกา ${winType}`];
        let coinDetailLines = [];
        
        active.forEach(p => {
            let diff = p.points - beforePoints[p.id];
            if (diff > 0) { coinDetailLines.push(`[PID:${p.id}] ได้รับ ${diff} เหรียญ`); } 
            else if (diff < 0) { losers.push(p.id); coinDetailLines.push(`[PID:${p.id}] เสีย ${Math.abs(diff)} เหรียญ`); } 
            else { coinDetailLines.push(`[PID:${p.id}] ไม่ได้ไม่เสียเหรียญ`); }
        });
        
        let resultDetail = cardDetailLines.join('<br>') + '<br><br>' + coinDetailLines.join('<br>');
        syncStateToAll(); broadcastTransition('จบเกมด้วยกติกาพิเศษ! กำลังสรุปผล...', 4000);
        
        setTimeout(() => {
            gameState.status = 'END'; syncStateToAll();
            let isFinalEnd = gameState.players.some(p => p.points === 0);
            let modalData = { type: 'SHOW_RESULT', title, detail: resultDetail, isEnd: isFinalEnd, winners, losers };
            if(myPeerId) window.showResultModal(title, resultDetail, isFinalEnd, winners, losers);
            guestConnections.forEach(c => {
                if (c && c.peer !== myPeerId) {
                    sendSoundMessageToConn(c, modalData);
                }
            });
        }, 4000);
    }, 4500);
}

function resetPlayerTimeoutOnValidHumanAction(p) {
    if (!p || p.isBot || p.temporaryBot || p.permanentBot) return;
    p.consecutiveTimeouts = 0;
    p.nextTurnTimeLimit = 40;
}

function processPlayerAction(pId, action, cardIndex, isInternalBotOrSystem = false) {
    let p = gameState.players.find(x => x.id === pId);
    if (!p || p.isOut) return;

    // ป้องกันผู้เล่นจริงส่งคำสั่งแทรกในขณะที่ตัวละครถูกควบคุมโดยบอท (ชั่วคราวหรือถาวร)
    if (!isInternalBotOrSystem && (p.isBot || p.temporaryBot || p.permanentBot)) {
        return;
    }

    if (action === 'SPECIAL_WIN') {
        if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(p);
        resolveSpecialWin(p.id, cardIndex);
        return;
    }

    if (gameState.status === 'PRE_GAME') {
        if (action === 'SKIP_PRE' && !p.isBot) {
            if (!gameState.skipPreVotes) gameState.skipPreVotes = [];
            if (!gameState.skipPreVotes.includes(pId)) {
                gameState.skipPreVotes.push(pId); syncStateToAll();
                let realPlayers = gameState.players.filter(x => !x.isOut && !x.isBot);
                if (gameState.skipPreVotes.length >= realPlayers.length) {
                    clearInterval(turnTimerInterval); triggerSound('60'); gameState.status = 'PLAYING';
                    let firstPlayer = gameState.players[gameState.turnIndex] || gameState.players[0];
                    restorePlayerIfReturnRequested(firstPlayer);
                    broadcastAnnounce(`📢 ผู้เล่นทุกคนพร้อมแล้ว! ขอเชิญ [PID:${firstPlayer.id}] จั่วไพ่ใบแรกและทิ้งลงกองเพื่อเริ่มเกมได้เลยครับ`);
                    syncStateToAll(); startTurnTimer(); checkBotTurn();
                }
            }
            return;
        }

        let handSum = p.hand.reduce((sum, c) => sum + c.val, 0);
        let isFirstPlayer = (pId === gameState.players[0].id);

        if (action === 'KANG_25' && handSum >= 25) {
            broadcastKangAnimation(p.id, 'cang25');
            setTimeout(() => resolveKang(p.id, false, false, true), 4500);
        } else if (action === 'KANG' && isFirstPlayer && handSum < 25) {
            broadcastKangAnimation(p.id, 'cang');
            setTimeout(() => resolveKang(p.id, false, false, false), 4500);
        }
        return;
    }

    if (action === 'FLOW') {
        if (gameState.status !== 'PLAYING') return;
        let flowPlayer = gameState.players.find(x => x.id === pId);
        if (!flowPlayer || flowPlayer.isOut) return;
        
        // --- การตั้งค่าป้องกันที่ 1 ---
        // ป้องกันการไหลถ้าผู้เล่นคนนั้นทำการจั่วไพ่แล้ว หรือทิ้งไพ่ปกติไปแล้วในเทิร์นตัวเอง
        if (flowPlayer.hasDrawnTurn || flowPlayer.hasDiscardedTurn) {
            return;
        }

        // --- เพิ่ม Logic Guard ป้องกันการไหลแซง Turn ---
        let currentPlayerId = gameState.players[gameState.turnIndex].id;
        if (pId !== currentPlayerId) {
            return; // ปฏิเสธ Action หากไม่ใช่ Turn ของผู้เล่นที่กดไหล
        }
        // ----------------------------------------------

        let topCard = discardPile[discardPile.length - 1]; let card = flowPlayer.hand[cardIndex];
        if (topCard && card && card.rank === topCard.rank) {
            if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(flowPlayer);
            triggerSound('follow'); let dropped = flowPlayer.hand.splice(cardIndex, 1)[0]; discardPile.push(dropped);
            flowPlayer.hasFlowedThisTurn = true; 

            let victimId = gameState.flowSourceId;
            let victim = gameState.players.find(v => v.id === victimId);
            if(!victim || victim.isOut) {
                let victimIndex = (gameState.turnIndex - 1 + gameState.players.length) % gameState.players.length;
                while(gameState.players[victimIndex].isOut) victimIndex = (victimIndex - 1 + gameState.players.length) % gameState.players.length;
                victim = gameState.players[victimIndex];
            }

            if (victim) {
                let flowMult = dropped.rank === 'A' ? 2 : 1; let penalty = 10 * flowMult;
                victim.points -= penalty; if(victim.points < 0) victim.points = 0; flowPlayer.points += penalty;
                broadcastAnnounce(`[PID:${flowPlayer.id}] ไหลไพ่ ${THAI_RANKS[dropped.rank]} ${THAI_SUITS[dropped.suit]}! ได้รับ ${penalty} เหรียญ จาก [PID:${victim.id}]`);
                broadcastFloatSC(victim.id, -penalty); broadcastFloatSC(flowPlayer.id, penalty);
            }

            gameState.topCardOwnerId = flowPlayer.id; 

            if(flowPlayer.hand.length === 0) { 
                broadcastKangAnimation(flowPlayer.id, 'knock');
                setTimeout(() => resolveKang(flowPlayer.id, false, 'FLOW_KNOCK'), 4500); 
            } 
            else {
                gameState.status = 'TRANSITION'; clearInterval(turnTimerInterval); syncStateToAll();
                broadcastTransition('ไหลไพ่สำเร็จ! รอสักครู่...', 2500);
                setTimeout(() => {
                    gameState.status = 'PLAYING';
                    syncStateToAll();
                    startTurnTimer();
                    if (flowPlayer.isBot) checkBotTurn();
                }, 2500);
            }
        }
        return;
    }

    let cp = gameState.players[gameState.turnIndex];
    if(cp.id !== pId || gameState.status !== 'PLAYING') return; 

    if(action === 'DRAW' && !cp.hasDrawnTurn && !cp.hasFlowedThisTurn) {
        if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(cp);
        if(deck.length > 0) {
            triggerSound('jua'); let drawnCard = deck.pop();
            cp.hand.push(drawnCard); cp.hasDrawnTurn = true;
            let publicMsg = `[PID:${cp.id}] จั่วไพ่ 1 ใบ`;
            if (cp.isBot && !cp.temporaryBot) { broadcastAnnounce(publicMsg); } 
            else {
                let privateMsg = `[PID:${cp.id}] จั่วได้ไพ่ ${THAI_RANKS[drawnCard.rank]} ${THAI_SUITS[drawnCard.suit]}`;
                if (myPeerId === cp.id) announce(privateMsg); else announce(publicMsg);
                guestConnections.forEach(c => {
                    if (c.peer === cp.id) c.send({ type: 'ANNOUNCE', msg: privateMsg }); else c.send({ type: 'ANNOUNCE', msg: publicMsg });
                });
            }
            syncStateToAll();
        } else resolveKang(cp.id, true); // กองไพ่หมดให้จบเกมทันที
    }
    else if(action === 'DISCARD') {
        let card = cp.hand[cardIndex];
        if (!card) return;
        if (!cp.hasDiscardedTurn && cp.hasDrawnTurn) {
            if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(cp);
            triggerSound('select'); let dropped = cp.hand.splice(cardIndex, 1)[0]; discardPile.push(dropped);
            
            let isFirstDiscardInTurn = !cp.hasDiscardedTurn;
            cp.hasDiscardedTurn = true; cp.discardedRankThisTurn = dropped.rank;
            
            if (isFirstDiscardInTurn) {
                gameState.flowSourceId = cp.id;
            }

            gameState.topCardOwnerId = cp.id;
            broadcastAnnounce(`[PID:${cp.id}] วางไพ่ ${THAI_RANKS[dropped.rank]} ${THAI_SUITS[dropped.suit]}`);
            if(cp.hand.length === 0) {
                broadcastKangAnimation(cp.id, 'knock');
                setTimeout(() => resolveKang(cp.id, false, 'DRAW_KNOCK'), 4500);
            } else syncStateToAll();
        } 
        else if (cp.hasDiscardedTurn && card.rank === cp.discardedRankThisTurn) {
            if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(cp);
            triggerSound('select'); let dropped = cp.hand.splice(cardIndex, 1)[0]; discardPile.push(dropped);
            broadcastAnnounce(`[PID:${cp.id}] วางไพ่ ${THAI_RANKS[dropped.rank]} ${THAI_SUITS[dropped.suit]} เพิ่มเติม (เลขเดียวกัน)`);
            if(cp.hand.length === 0) {
                broadcastKangAnimation(cp.id, 'knock');
                setTimeout(() => resolveKang(cp.id, false, 'DRAW_KNOCK'), 4500);
            } else syncStateToAll();
        }
    }
    else if(action === 'END_TURN' && (cp.hasDiscardedTurn || cp.hasFlowedThisTurn)) {
        if (isEndingTurn) return;
        isEndingTurn = true;
        if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(cp);
        cp.hasFinishedFirstTurn = true;
        cp.turnCount = (cp.turnCount || 0) + 1;
        triggerSound('select'); broadcastAnnounce(`[PID:${cp.id}] กดยืนยันจบตา`);
        gameState.status = 'TRANSITION'; clearInterval(turnTimerInterval); syncStateToAll();
        broadcastTransition('กำลังเปลี่ยนตาถัดไป...', 2000);
        setTimeout(() => { gameState.status = 'PLAYING'; nextTurn(); }, 2000);
    }
    else if(action === 'KANG' && !cp.hasDrawnTurn && !cp.hasFlowedThisTurn && discardPile.length > 0) {
        if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(cp);
        broadcastKangAnimation(cp.id, 'cang'); setTimeout(() => resolveKang(cp.id), 4500);
    }
}

function resolveKang(callerId, isDeckEmpty = false, winReason = false, isKang25 = false) {
    window.nextRoundHostId = callerId;
    gameState.status = 'TRANSITION'; clearInterval(turnTimerInterval);
    let active = gameState.players.filter(p => !p.isOut);
    let totalPlayersInRoom = active.length;
    let beforePoints = {}; active.forEach(p => beforePoints[p.id] = p.points);

    active.forEach(p => { 
        p.handSum = p.hand.reduce((s, c) => s + c.val, 0); 
        p.aceCount = p.hand.filter(c => c.rank === 'A').length;
    });
    
    let caller = active.find(p => p.id === callerId); 
    let resultDetail = "", title = "", winnerColorName = "";
    const getBasePay = (player) => 10 + (player.aceCount * 10);

    if (winReason === 'FLOW_KNOCK' || winReason === 'DRAW_KNOCK' || winReason === true) {
        let basePay = (winReason === 'DRAW_KNOCK') ? 20 : 10;
        let reasonStr = (winReason === 'DRAW_KNOCK') ? 'จั่วแล้วทิ้งไพ่หมดมือ' : 'ไหลไพ่หมดมือน็อก';
        if (winReason === true) reasonStr = 'น็อค!';
        
        winnerColorName = caller.colorInfo.name;
        title = `[PID:${caller.id}] ${reasonStr}! (ชนะรับ ${basePay} เหรียญ จากทุกคน)`;
        active.forEach(p => { 
            if(p.id !== caller.id) { p.points -= basePay; if(p.points < 0) p.points = 0; caller.points += basePay; } 
        });
    } else if (isDeckEmpty) {
        let lowest = active.reduce((min, p) => p.handSum < min.handSum ? p : min, active[0]);
        let basePay = getBasePay(lowest);
        winnerColorName = lowest.colorInfo.name;
        title = `กองไพ่หมด! [PID:${lowest.id}] แต้มต่ำสุด ${lowest.handSum} แต้ม (ชนะ)`;
        active.forEach(p => { 
            if(p.id !== lowest.id) { p.points -= basePay; if(p.points < 0) p.points = 0; lowest.points += basePay; } 
        });
    } else { 
        let others = active.filter(p => p.id !== caller.id);
        let lowestOther = others.reduce((min, p) => p.handSum < min.handSum ? p : min, others[0]);

        if (caller.handSum <= lowestOther.handSum) { // เปลี่ยนจาก < เป็น <= เพื่อให้ชนะเสมอเมื่อแต้มต่ำสุดเท่ากัน
            winnerColorName = caller.colorInfo.name;
            let payAmount = getBasePay(caller); if (isKang25) payAmount = payAmount * 2; 
            title = `[PID:${caller.id}] ${isKang25?'แคง 25':'แคง'} สำเร็จ!`;
            others.forEach(p => { p.points -= payAmount; if(p.points < 0) p.points = 0; caller.points += payAmount; });
        } else {
            winnerColorName = lowestOther.colorInfo.name;
            let payAmount = getBasePay(lowestOther) * totalPlayersInRoom; 
            title = `[PID:${caller.id}] ${isKang25?'แคง 25':'แคง'} แหก! ([PID:${lowestOther.id}] ชนะ)`;
            caller.points -= payAmount; if(caller.points < 0) caller.points = 0; lowestOther.points += payAmount;
        }
    }

    let winners = []; let losers = [];
    let cardDetailLines = []; let coinDetailLines = [];

    active.forEach(p => {
        let diff = p.points - beforePoints[p.id];
        if (diff > 0) { winners.push(p.id); coinDetailLines.push(`[PID:${p.id}] ได้รับ ${diff} เหรียญ`); } 
        else if (diff < 0) { losers.push(p.id); coinDetailLines.push(`[PID:${p.id}] เสีย ${Math.abs(diff)} เหรียญ`); } 
        else { coinDetailLines.push(`[PID:${p.id}] ไม่ได้ไม่เสียเหรียญ`); }
        cardDetailLines.push(`[PID:${p.id}] เปิดไพ่มาแล้วได้แต้มรวม ${p.handSum} แต้ม`);
    });

    resultDetail = cardDetailLines.join('<br>') + '<br><br>' + coinDetailLines.join('<br>');
    syncStateToAll(); broadcastTransition('จบเกม! กำลังสรุปผล...', 4000);

    setTimeout(() => {
        gameState.status = 'END'; syncStateToAll();
        let isFinalEnd = gameState.players.some(p => p.points === 0);
        let modalData = { type: 'SHOW_RESULT', title, detail: resultDetail, isEnd: isFinalEnd, winners, losers };
        if(myPeerId) window.showResultModal(title, resultDetail, isFinalEnd, winners, losers);
        guestConnections.forEach(c => {
            if (c && c.peer !== myPeerId) {
                sendSoundMessageToConn(c, modalData);
            }
        });
    }, 4000);
}

window.closeResultModal = () => {
    if (window.isStartingRound) return;
    window.isStartingRound = true;
    let btn = document.getElementById('btn-next-round');
    if (btn) { btn.disabled = true; btn.style.display = 'none'; }
    document.getElementById('result-modal').style.display = 'none';
    updateHomeBtnVisibility();
    if(isHost) { guestConnections.forEach(c => sendSoundMessageToConn(c, { type: 'START_COUNTDOWN' })); doCountdownAndStart(); }
};

// =========================================================
// Client Rendering Logic
// =========================================================
function renderClientGame(publicState, privateState) {
    window.currentPublicState = publicState; // [แคช state ล่าสุดไว้ใช้ดึงเวลาเสกไพ่ใน Client]

    if (publicState.status === 'PRE_GAME' || publicState.status === 'PLAYING') { document.getElementById('result-modal').style.display = 'none'; }
    updateHomeBtnVisibility();

    localPlayerState = privateState;
    const turnId = publicState.turnId;
    const isFirstPlayer = (myPeerId === publicState.playersInfo[0].id);
    const myPublicInfo = publicState.playersInfo.find(p => p.id === myPeerId);
    const isControlledByBot = !!(localPlayerState.temporaryBot || (myPublicInfo && (myPublicInfo.isBot || myPublicInfo.temporaryBot)));
    const isMyTurn = (turnId === myPeerId) && !isControlledByBot;
    
    publicState.playersInfo.forEach(p => {
        let pName = (p.name && String(p.name).trim() !== '') ? String(p.name).trim() : (globalPlayersMap[p.id] && globalPlayersMap[p.id].name ? globalPlayersMap[p.id].name : 'ผู้เล่น');
        globalPlayersMap[p.id] = { isBot: p.isBot, temporaryBot: p.temporaryBot, permanentBot: p.permanentBot, name: pName, colorName: p.colorName, colorHex: p.colorHex };
    });
    let currPlayerDisp = resolveName(turnId);
    
    const turnInd = document.getElementById('turn-indicator');
    const timerBarContainer = document.getElementById('timer-bar-container');
    
    if (publicState.status === 'END') {
        if(timerBarContainer) timerBarContainer.style.display = 'none';
        turnInd.style.display = 'none'; 
        document.getElementById('main-table-area').style.display = 'none';
        document.getElementById('action-bar-container').style.display = 'none';
    } else {
        turnInd.style.display = 'block'; 
        if(timerBarContainer) timerBarContainer.style.display = 'block';
        document.getElementById('main-table-area').style.display = 'flex';
        document.getElementById('action-bar-container').style.display = 'flex';
        if (publicState.status === 'PRE_GAME') {
            turnInd.innerText = "เตรียมตัว! เช็คไพ่ในมือ (40 วินาที)";
        } else {
            let oldTurnTxt = turnInd.innerText;
            let newTurnTxt = (turnId === myPeerId && !isControlledByBot) ? `🔥 ถึงรอบของ ${currPlayerDisp} แล้ว!` : `ถึงรอบของ ${currPlayerDisp} แล้ว`;
            if (oldTurnTxt !== newTurnTxt && publicState.status === 'PLAYING') {
                turnInd.innerText = newTurnTxt;
                let topC = publicState.topDiscard;
                let topTxt = topC ? `${THAI_RANKS[topC.rank]} ${THAI_SUITS[topC.suit]}` : 'ยังไม่มีไพ่บนกอง';
                if(turnId === myPeerId && !isControlledByBot) announce(`ถึงรอบของ ${currPlayerDisp} แล้ว กรุณาลงมือเล่น (ไพ่บนกองคือ ${topTxt})`);
                else announce(`ถึงรอบของ ${currPlayerDisp} แล้ว (ไพ่บนกองคือ ${topTxt})`, false);
            }
        }
    }

    const renderStatusBar = (containerId) => {
        const container = document.getElementById(containerId);
        if (!container) return;
        container.innerHTML = '';
        let groupAriaLabel = "";

        publicState.playersInfo.forEach((p) => {
            if (!p.isOut) {
                let dispName = resolveName(p.id);
                groupAriaLabel += `${dispName} มี ${p.points} เหรียญ. `;
                let item = document.createElement('div');
                item.className = `status-avatar-item ${p.id === turnId && publicState.status === 'PLAYING' ? 'active-turn' : ''}`;
                item.id = `avatar-${containerId}-${p.id}`;
                item.style.borderColor = p.colorHex;
                item.innerHTML = `
                    <div class="status-coin-header" aria-hidden="true">${p.points} เหรียญ</div>
                    <div class="status-avatar-icon" style="background-color:${p.colorHex}; color:#fff;" aria-hidden="true">👤</div>
                    <div class="status-player-label" style="color:${p.colorHex};" aria-hidden="true">${dispName}</div>
                `;
                container.appendChild(item);
            }
        });
        container.setAttribute('tabindex', '0');
        container.setAttribute('role', 'group');
        container.setAttribute('aria-label', `สถานะผู้เล่นทั้งหมด: ${groupAriaLabel}`);
    };

    renderStatusBar('players-status-bar'); renderStatusBar('result-status-bar');

    const btnDraw = document.getElementById('btn-draw'); const btnKang = document.getElementById('btn-kang');
    const btnKang25 = document.getElementById('btn-kang25'); const btnEnd = document.getElementById('btn-end-turn');
    const btnSkip = document.getElementById('btn-skip-pre'); const btnSpecial = document.getElementById('btn-special');
    let topC = publicState.topDiscard;

    if (topC) currentTopCardText = `ไพ่บนกองทิ้งปัจจุบันคือ ${THAI_RANKS[topC.rank]} ${THAI_SUITS[topC.suit]}`;
    else currentTopCardText = `ยังไม่มีไพ่บนกองทิ้ง`;

    let hasVoted = (publicState.skipVotes && publicState.skipVotes.includes(myPeerId)) || isControlledByBot;

    if (publicState.status === 'PRE_GAME') {
        let myHandSum = (localPlayerState.hand || []).reduce((s, c) => s + (c.val || getCardValue(c.rank)), 0);
        btnDraw.style.display = 'none'; btnEnd.style.display = 'none';
        
        if (isFirstPlayer) {
            if (myHandSum < 25) {
                btnKang.style.display = 'inline-block'; btnKang.disabled = hasVoted; 
                btnKang25.style.display = 'none'; btnKang25.disabled = true;
            } else {
                btnKang.style.display = 'none'; btnKang.disabled = true;
                btnKang25.style.display = 'inline-block'; btnKang25.disabled = hasVoted; 
            }
        } else {
            btnKang.style.display = 'none'; btnKang.disabled = true;
            if (myHandSum >= 25) { btnKang25.style.display = 'inline-block'; btnKang25.disabled = hasVoted; } 
            else { btnKang25.style.display = 'none'; btnKang25.disabled = true; }
        }
        btnSkip.style.display = 'inline-block'; btnSkip.disabled = hasVoted;

    } else if (publicState.status === 'PLAYING' || publicState.status === 'TRANSITION') {
        btnKang25.style.display = 'none'; btnSkip.style.display = 'none';
        btnDraw.style.display = 'inline-block'; btnEnd.style.display = 'inline-block'; btnKang.style.display = 'inline-block';

        if(isMyTurn && publicState.status === 'PLAYING' && !isControlledByBot) {
            btnDraw.disabled = localPlayerState.hasDrawn || localPlayerState.hasFlowedThisTurn;
            btnKang.disabled = localPlayerState.hasDrawn || localPlayerState.hasFlowedThisTurn || publicState.discardPileCount === 0;
            btnEnd.disabled = !(localPlayerState.hasDiscarded || localPlayerState.hasFlowedThisTurn);
        } else { btnDraw.disabled = true; btnKang.disabled = true; btnEnd.disabled = true; }
    }

    currentSpecialWinType = getSpecialWinType(localPlayerState.hand, localPlayerState);

    if (currentSpecialWinType && (publicState.status === 'PRE_GAME' || publicState.status === 'PLAYING') && !isControlledByBot) {
        btnSpecial.style.display = 'inline-block'; btnSpecial.innerText = currentSpecialWinType;
        btnSpecial.disabled = false;
    } else {
        btnSpecial.style.display = 'none';
        btnSpecial.disabled = true;
    }

const deckCountTxt = document.getElementById('deck-count-text');

if (deckCountTxt) {
    deckCountTxt.innerText = publicState.deckCount;
}

const deckAccessibleObject = document.getElementById('deck-accessible-object');

if (deckAccessibleObject) {
    deckAccessibleObject.setAttribute(
        'aria-label',
        `กองไพ่สำหรับจั่ว เหลือ ${publicState.deckCount} ใบ`
    );
}
    let dp = document.getElementById('discard-pile');
    
    // --- [เพิ่มมิติการแสดงผล] Visual Trigger: อัปเดตกองทิ้ง ---
    if(topC) {
        let currentTopId = `${topC.rank}-${topC.suit}`;
        if (window._lastTopCardId !== currentTopId) {
            window._lastTopCardId = currentTopId;
            dp.style.animation = 'none';
            void dp.offsetWidth; // trigger reflow
            dp.style.animation = 'cardDropScale 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275)';
        }

        let tRank = THAI_RANKS[topC.rank]; let tSuit = THAI_SUITS[topC.suit];
        dp.innerHTML = `<div class="card ${['♥','♦'].includes(topC.suit)?'red':'black'}" aria-label="ไพ่กองทิ้งใบบนสุดคือ ${tRank} ${tSuit}" role="img">
                           <div aria-hidden="true">${topC.rank}</div><div class="card-suit" aria-hidden="true">${topC.suit}</div>
                        </div>`;
    } else {
        window._lastTopCardId = null;
        dp.innerHTML = '<div style="color:#a0a0a0; font-size:0.85rem;">กองทิ้งว่างเปล่า</div>'; 
        dp.setAttribute('aria-label', 'กองไพ่ทิ้งยังว่างเปล่า');
    }

    const handUi = document.getElementById('my-hand-ui'); handUi.innerHTML = '';
    if (publicState.status !== 'END') {
        
        // --- การตั้งค่าป้องกันที่ 2 ---
        // ปรับเงื่อนไข canFlowCard ไม่ให้เกิดสถานะการไหลได้ ถ้าตัวผู้เล่นจั่วไพ่แล้ว หรือทิ้งไพ่รอบนี้ไปแล้ว
        (localPlayerState.hand || []).forEach((c, index) => {
            let cardEl = document.createElement('button');
            cardEl.className = `card ${['♥','♦'].includes(c.suit)?'red':'black'}`;
            
            // --- [เพิ่มมิติการแสดงผล] Visual Trigger: ไพ่ทยอยแจกเข้ามือ ---
            cardEl.style.animation = `cardSlideUp 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275) ${index * 0.08}s both`;

            let tRank = THAI_RANKS[c.rank]; let tSuit = THAI_SUITS[c.suit];
            let actionHint = ""; cardEl.disabled = true;
            
            let canFlowCard = (topC && c.rank === topC.rank && !localPlayerState.hasDrawn && !localPlayerState.hasDiscarded);

            if (publicState.status === 'PLAYING' && !isControlledByBot) {
                if (canFlowCard) {
                    if (isMyTurn) { // เป็น Turn ตัวเองเท่านั้นจึงจะไหลได้
                        cardEl.disabled = false; actionHint = " กดเพื่อไหลไพ่ใบนี้";
                        cardEl.style.boxShadow = "0 0 12px var(--blue)"; 
                        cardEl.onclick = () => window.clientAction('FLOW', index);
                    } else { // ไม่ใช่ Turn ตัวเอง ไม่สามารถกดได้ และทำให้แสดงสถานะ dim/disabled
                        cardEl.disabled = true;
                        cardEl.style.opacity = '0.5';
                        cardEl.style.cursor = 'not-allowed';
                    }
                }
                else if (isMyTurn) {
                    if (!localPlayerState.hasDiscarded) {
                        if (localPlayerState.hasDrawn && !localPlayerState.hasFlowedThisTurn) {
                            cardEl.disabled = false; actionHint = " กดเพื่อทิ้งไพ่ใบนี้";
                            cardEl.onclick = () => window.clientAction('DISCARD', index);
                        }
                    } else if (c.rank === localPlayerState.discardedRank) {
                        cardEl.disabled = false; actionHint = " กดเพื่อทิ้งเพิ่มเติม (เลขเดียวกัน)";
                        cardEl.style.boxShadow = "0 0 12px var(--green)";
                        cardEl.onclick = () => window.clientAction('DISCARD', index);
                    }
                }
            }
            
            cardEl.setAttribute('aria-label', `ไพ่ ${tRank} ${tSuit}${actionHint}`);
            cardEl.innerHTML = `<div aria-hidden="true">${c.rank}</div><div class="card-suit" aria-hidden="true">${c.suit}</div>`;
            handUi.appendChild(cardEl);
        });
    }
}

window.clientAction = (action, index = 0) => {
    const myPub = window.currentPublicState && window.currentPublicState.playersInfo
        ? window.currentPublicState.playersInfo.find(p => p.id === myPeerId)
        : null;
    if (localPlayerState.temporaryBot || (myPub && (myPub.isBot || myPub.temporaryBot))) {
        return;
    }

    // --- [เพิ่มมิติการแสดงผล] Visual Trigger: แอนิเมชันเมื่อผู้เล่นกดทำ Action ส่งข้อมูล ---
    try {
        if (action === 'DISCARD' || action === 'FLOW') {
            const handUi = document.getElementById('my-hand-ui');
            const cardEl = handUi ? handUi.children[index] : null;
            const discardPileEl = document.getElementById('discard-pile');
            
            if (cardEl && discardPileEl) {
                const cardRect = cardEl.getBoundingClientRect();
                const targetRect = discardPileEl.getBoundingClientRect();
                
                const clone = cardEl.cloneNode(true);
                clone.style.position = 'fixed';
                clone.style.left = `${cardRect.left}px`;
                clone.style.top = `${cardRect.top}px`;
                clone.style.width = `${cardRect.width}px`;
                clone.style.height = `${cardRect.height}px`;
                clone.style.zIndex = '9999';
                clone.style.transition = 'all 0.4s cubic-bezier(0.25, 1, 0.5, 1)';
                clone.style.pointerEvents = 'none';
                clone.style.margin = '0';
                document.body.appendChild(clone);
                
                cardEl.style.opacity = '0'; // ซ่อนไพ่บนมือทันทีเพื่อความสมูท
                
                setTimeout(() => {
                    clone.style.left = `${targetRect.left + (targetRect.width/2) - (cardRect.width/2)}px`;
                    clone.style.top = `${targetRect.top + (targetRect.height/2) - (cardRect.height/2)}px`;
                    clone.style.transform = `scale(0.8) rotate(${Math.random() * 30 - 15}deg)`;
                    clone.style.opacity = '0.7';
                }, 10);
                
                setTimeout(() => clone.remove(), 400);
            }
        } else if (action === 'DRAW') {
            const deckEl = document.getElementById('deck-group-box');
            if (deckEl) {
                const deckRect = deckEl.getBoundingClientRect();
                const clone = document.createElement('div');
                clone.className = 'card';
                // สร้างจำลองหลังไพ่คร่าวๆ สำหรับให้ลอยเข้ามือ
                clone.style.background = 'linear-gradient(135deg, var(--blue, #1e3c72), #2a5298)';
                clone.style.border = '2px solid #fff';
                clone.style.position = 'fixed';
                clone.style.left = `${deckRect.left}px`;
                clone.style.top = `${deckRect.top}px`;
                clone.style.width = '60px'; 
                clone.style.height = '90px';
                clone.style.borderRadius = '8px';
                clone.style.zIndex = '9999';
                clone.style.transition = 'all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275)';
                clone.style.pointerEvents = 'none';
                document.body.appendChild(clone);
                
                setTimeout(() => {
                    clone.style.left = `50%`;
                    clone.style.top = `90%`; // ลอยมาทางผู้เล่นด้านล่างจอ
                    clone.style.transform = 'translate(-50%, -50%) scale(1.2)';
                    clone.style.opacity = '0';
                }, 10);
                
                setTimeout(() => clone.remove(), 400);
            }
        }
    } catch(e) { console.warn('Animation error ignored', e); }
    // -----------------------------------------------------------

    if(isHost) processPlayerAction(myPeerId, action, index);
    else hostConnection.send({ type: 'PLAYER_ACTION', playerId: myPeerId, action, cardIndex: index });
};

window.showFloatingScore = (targetId, amt) => {
    let target = document.getElementById(`avatar-players-status-bar-${targetId}`);
    if (target) {
        let floatEl = document.createElement('div'); 
        floatEl.className = 'floating-sc';
        
        // --- [เพิ่มมิติการแสดงผล] Visual Trigger: แอนิเมชันเหรียญที่เด้งขึ้น ---
        floatEl.style.transition = 'all 1.5s ease-out';
        floatEl.style.position = 'absolute';
        floatEl.style.top = '0px';
        floatEl.style.left = '50%';
        floatEl.style.transform = 'translate(-50%, 0) scale(0.5)';
        floatEl.style.opacity = '1';
        floatEl.style.fontWeight = 'bold';
        floatEl.style.fontSize = '1.2rem';
        floatEl.style.textShadow = '0 0 5px rgba(0,0,0,0.8)';
        
        floatEl.style.color = amt > 0 ? '#32cd32' : '#ff4d4d'; 
        floatEl.innerText = amt > 0 ? `+${amt}` : amt; 
        
        floatEl.setAttribute('aria-hidden', 'true'); 
        target.style.position = 'relative'; // Ensure relative positioning
        target.appendChild(floatEl); 
        
        // Trigger reflow & Animate
        void floatEl.offsetWidth; 
        setTimeout(() => {
            floatEl.style.transform = 'translate(-50%, -40px) scale(1.2)';
            floatEl.style.opacity = '0';
        }, 10);

        setTimeout(() => floatEl.remove(), 1500);
    }
};

window.showResultModal = (title, detail, isEnd, winners, losers) => {
    const srOnlyAnnouncer = document.getElementById('sr-only-announcer'); if (srOnlyAnnouncer) srOnlyAnnouncer.innerText = '';
    const visAnnouncer = document.getElementById('visible-game-announcer'); if (visAnnouncer) visAnnouncer.innerText = '';
    window.isStartingRound = false;

    let resolvedTitle = title.replace(/\[PID:(.*?)\]/g, (match, id) => resolveName(id));
    let resolvedDetail = detail.replace(/\[PID:(.*?)\]/g, (match, id) => resolveName(id));

    document.getElementById('result-title').innerText = resolvedTitle;
    document.getElementById('result-details').innerHTML = resolvedDetail;
    
    document.getElementById('result-modal').style.display = 'flex';
    updateHomeBtnVisibility();

    if (winners && losers) {
        winners.forEach(id => { let el = document.getElementById(`avatar-result-status-bar-${id}`); if (el) el.classList.add('anim-happy'); });
        losers.forEach(id => { let el = document.getElementById(`avatar-result-status-bar-${id}`); if (el) el.classList.add('anim-sad'); });
    }

    let btnNext = document.getElementById('btn-next-round'); let waitTxt = document.getElementById('guest-waiting-next-round');
    
    let hasZeroCoins = (gameState.players && gameState.players.some(p => p.points === 0)) ||
                       (window.currentPublicState && window.currentPublicState.playersInfo && window.currentPublicState.playersInfo.some(p => p.points === 0));

    if (btnNext) {
        if (isHost && hasZeroCoins) {
            btnNext.style.display = 'none';
            if (waitTxt) waitTxt.style.display = 'none';
        } else if (isEnd) {
            btnNext.style.display = 'block'; btnNext.disabled = false; btnNext.innerText = "กลับหน้าหลัก"; btnNext.onclick = () => window.location.reload();
            if (waitTxt) waitTxt.style.display = 'none';
        } else {
            if (isHost) {
                btnNext.style.display = 'block'; btnNext.disabled = false; btnNext.innerText = "เริ่มรอบใหม่"; btnNext.onclick = closeResultModal;
                if (waitTxt) waitTxt.style.display = 'none';
            } else {
                btnNext.style.display = 'none';
                if (waitTxt) waitTxt.style.display = 'block';
            }
        }
    }

    if(myPeerId) {
        if (winners && winners.includes(myPeerId)) playSound('win');
        else if (losers && losers.includes(myPeerId)) playSound('lost');
        else playSound('no'); 
    }

    setTimeout(() => {
        if (!isLocalTempBotModalOpen) {
            const resultTitle = document.getElementById('result-title');
            if (resultTitle) resultTitle.focus();
        }
    }, 200);
};

// =========================================================
// --- CHEAT MODE LOGIC (รวมฟังก์ชันไว้ล่างสุด) ---
// =========================================================
function setCheatHand(type) {
    let newHand = [];
    if (type === "ตอง") {
        newHand = [
            { suit: '♠', rank: '7', value: 5, val: 5, id: '7♠' },
            { suit: '♥', rank: '7', value: 5, val: 5, id: '7♥' },
            { suit: '♦', rank: '7', value: 5, val: 5, id: '7♦' },
            { suit: '♣', rank: '7', value: 5, val: 5, id: '7♣' },
            { suit: '♠', rank: 'K', value: 10, val: 10, id: 'K♠' },
            { suit: '♥', rank: 'K', value: 10, val: 10, id: 'K♥' },
            { suit: '♦', rank: 'K', value: 10, val: 10, id: 'K♦' }
        ];
    } else if (type === "ดอก" || type === "สี") {
        newHand = [
            { suit: '♠', rank: '2', value: 5, val: 5, id: '2♠' },
            { suit: '♠', rank: '3', value: 5, val: 5, id: '3♠' },
            { suit: '♠', rank: '4', value: 5, val: 5, id: '4♠' },
            { suit: '♠', rank: '5', value: 5, val: 5, id: '5♠' },
            { suit: '♠', rank: '6', value: 5, val: 5, id: '6♠' },
            { suit: '♠', rank: '7', value: 5, val: 5, id: '7♠' },
            { suit: '♠', rank: '8', value: 5, val: 5, id: '8♠' }
        ];
    } else if (type === "เรียง") {
        newHand = [
            { suit: '♥', rank: '4', value: 5, val: 5, id: '4♥' },
            { suit: '♥', rank: '5', value: 5, val: 5, id: '5♥' },
            { suit: '♥', rank: '6', value: 5, val: 5, id: '6♥' },
            { suit: '♣', rank: '8', value: 5, val: 5, id: '8♣' },
            { suit: '♣', rank: '9', value: 5, val: 5, id: '9♣' },
            { suit: '♣', rank: '10', value: 10, val: 10, id: '10♣' },
            { suit: '♣', rank: 'J', value: 10, val: 10, id: 'J♣' }
        ];
    } else if (type === "50") {
        newHand = [
            { suit: '♠', rank: '2', value: 50, val: 50, id: '2♠' },
            { suit: '♣', rank: 'Q', value: 50, val: 50, id: 'Q♣' },
            { suit: '♠', rank: 'A', value: 15, val: 15, id: 'A♠' },
            { suit: '♥', rank: 'A', value: 15, val: 15, id: 'A♥' },
            { suit: '♦', rank: 'A', value: 15, val: 15, id: 'A♦' },
            { suit: '♦', rank: '10', value: 10, val: 10, id: '10♦' },
            { suit: '♥', rank: 'K', value: 10, val: 10, id: 'K♥' }
        ];
    } else if (type === "3a") {
        newHand = [
            { suit: '♠', rank: 'A', value: 15, val: 15, id: 'A♠' },
            { suit: '♥', rank: 'A', value: 15, val: 15, id: 'A♥' },
            { suit: '♦', rank: 'A', value: 15, val: 15, id: 'A♦' },
            { suit: '♠', rank: '2', value: 50, val: 50, id: '2♠' },
            { suit: '♣', rank: 'Q', value: 50, val: 50, id: 'Q♣' },
            { suit: '♦', rank: 'K', value: 10, val: 10, id: 'K♦' },
            { suit: '♣', rank: 'K', value: 10, val: 10, id: 'K♣' }
        ];
    }

    if (isHost) {
        let p = gameState.players.find(x => x.id === myPeerId);
        if (p) {
            p.hand = newHand;
            syncStateToAll();
            announce("เสกไพ่สำเร็จ (Host)", false);
        }
    } else {
        localPlayerState.hand = newHand;
        if (window.currentPublicState) {
            renderClientGame(window.currentPublicState, localPlayerState);
        }
        announce("เสกไพ่สำเร็จ (Client)", false);
    }
}