import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
import { getDatabase, ref, set, onValue, update, runTransaction, remove, onDisconnect } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-database.js";

// Exact Firebase Configuration from original
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

// Car Avatars Pool
const CAR_POOL = [
    { icon: '🚗', name: 'รถเก๋งแดง' }, { icon: '🚕', name: 'แท็กซี่' },
    { icon: '🚙', name: 'รถจี๊ป' }, { icon: '🚌', name: 'รถบัส' },
    { icon: '🚎', name: 'รถทัวร์' }, { icon: '🏎️', name: 'รถแข่ง' },
    { icon: '🚓', name: 'รถตำรวจ' }, { icon: '🚑', name: 'รถพยาบาล' },
    { icon: '🚒', name: 'รถดับเพลิง' }, { icon: '🚐', name: 'รถตู้' },
    { icon: '🛻', name: 'กระบะ' }, { icon: '🚚', name: 'รถบรรทุก' }
];

const REST_TYPES = [
    { name: "คาเฟ่ริมทาง", icon: "☕" }, { name: "โรงแรมจิ้งหรีด", icon: "🏨" },
    { name: "ห้องน้ำกลางทุ่ง", icon: "🚽" }, { name: "ร้านข้าวแกง 24 ชั่วโมง", icon: "🍛" },
    { name: "ร้านตัดผมสุดสยอง", icon: "💈" }, { name: "Car care สกปรก", icon: "🧽" },
    { name: "อู่ซ่อมรถร้าง", icon: "🛠️" }, { name: "โรงน้ำชา", icon: "🍵" }
];

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
let autoAnswerTimer = null;
let isStartingGame = false;
let isShowingWinnerScene = false;
let myLastCorrectAnswerText = '';

// Queue สำหรับจัดการลำดับเสียงและ action ป้องกันการทับซ้อน
let isProcessingAction = false;
let actionQueue = [];

const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
const unlockAudio = () => { if (audioCtx.state === 'suspended') audioCtx.resume(); };
document.addEventListener('click', unlockAudio, { capture: true });
document.addEventListener('touchstart', unlockAudio, { capture: true });
document.addEventListener('keydown', unlockAudio, { capture: true });

let bgmSource = null;
let audioBuffers = {};

async function loadAudio(name) {
    if (audioBuffers[name]) return audioBuffers[name];
    try {
        const response = await fetch(`audio/${name}`);
        const arrayBuffer = await response.arrayBuffer();
        const audioBuffer = await audioCtx.decodeAudioData(arrayBuffer);
        audioBuffers[name] = audioBuffer;
        return audioBuffer;
    } catch (e) {
        return null;
    }
}

function playAudio(name, loop = false) {
    return new Promise(async (resolve) => {
        if (audioCtx.state === 'suspended') await audioCtx.resume();
        const buffer = await loadAudio(name);
        if (!buffer) { resolve(); return; }
        const source = audioCtx.createBufferSource();
        source.buffer = buffer;
        source.connect(audioCtx.destination);
        source.loop = loop;
        source.onended = () => resolve(source);
        source.start(0);
        if (loop && name === 'bgm.mp3') {
            if (bgmSource) { try { bgmSource.stop(); } catch(e){} }
            bgmSource = source;
        }
    });
}

async function playAudioSequence(list) {
    for (const name of list) { await playAudio(name); }
}

function stopBGM() {
    if (bgmSource) { try { bgmSource.stop(); } catch(e){} bgmSource = null; }
}

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

async function processActionQueue() {
    if (isProcessingAction || actionQueue.length === 0) return;
    isProcessingAction = true;
    const action = actionQueue.shift();
    
    // ประกาศ Screen Reader ทันทีเพื่อความสอดคล้องกับจังหวะเริ่มเหตุการณ์โดยไม่ต้องรอเสียงจบ
    announceSR(action.msg, 'polite');
    
    if (action.msg.includes('ทอยลูกเต๋าได้')) {
        const match = action.msg.match(/ทอยลูกเต๋าได้\s*(\d+)/);
        if (match) {
            const diceEl = document.getElementById('dice-visual');
            if (diceEl) {
                diceEl.textContent = match[1];
                diceEl.setAttribute('aria-hidden', 'true');
            }
        }
    }
    
    // รอจนกว่าเสียงที่เกี่ยวข้องกับเหตุการณ์นี้เล่นจบก่อน เพื่อรักษาลำดับเหตุการณ์ถัดไปไม่ให้ทับซ้อน
    if (action.audioKeys && action.audioKeys.length > 0) {
        await playAudioSequence(action.audioKeys);
    }
    
    isProcessingAction = false;
    processActionQueue();
}

function delayAsync(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }

window.switchScreen = function(screenId, focusHeadingId = null) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active-screen'));
    const target = document.getElementById(screenId);
    if (target) {
        target.classList.add('active-screen');
        if (focusHeadingId) {
            const h = document.getElementById(focusHeadingId);
            if (h) {
                h.setAttribute('tabindex', '-1');
                h.focus();
                const cleanupFocus = () => { h.removeAttribute('tabindex'); h.removeEventListener('blur', cleanupFocus); };
                h.addEventListener('blur', cleanupFocus);
            }
        }
    }
};

window.confirmNameAndEnterLobby = function() {
    playAudio('select.mp3');
    const nameInput = document.getElementById('player-name-input');
    myPlayerName = nameInput.value.trim() || 'นักแข่งใหม่';
    switchScreen('screen-lobby', 'lobby-heading');
    announceSR(`ยินดีต้อนรับ ${myPlayerName} เข้าสู่ล็อบบี้แรลลี่`);
};

window.toggleManual = function() {
    playAudio('select.mp3');
    const m = document.getElementById('manual-section');
    if (m.style.display === 'none') {
        m.style.display = 'block'; m.focus(); announceSR("เปิดคู่มือการเล่นแล้ว");
    } else {
        m.style.display = 'none'; announceSR("ปิดคู่มือการเล่นแล้ว");
    }
};

function initRoomListListener() {
    const roomsRef = ref(db, 'games/RallyThai/rooms');
    roomListListener = onValue(roomsRef, (snapshot) => {
        const rooms = snapshot.val() || {};
        const container = document.getElementById('room-list');
        if (!container || myPlayerId) return;
        container.innerHTML = '';
        let count = 0;
        for (const [rId, rData] of Object.entries(rooms)) {
            if (rData.status === 'waiting') {
                count++;
                const item = document.createElement('div');
                item.className = 'participant-card';
                item.style.cursor = 'pointer';
                const pCount = rData.players ? Object.keys(rData.players).length : 0;
                item.innerHTML = `<span><strong>ห้อง ${rId}</strong> (${pCount}/6 คน)</span>`;
                item.onclick = () => window.joinRallyRoom(rId);
                container.appendChild(item);
            }
        }
        if (count === 0) container.innerHTML = '<div style="color:#bdc3c7; padding:10px;">ไม่มีห้องที่กำลังรอผู้เล่นอยู่...</div>';
    });
}

function generateRallyBoard() {
    const questionsData = window.rallyQuestionsData || {};
    let allProvinces = Object.keys(questionsData);
    
    // Shuffle provinces
    for (let i = allProvinces.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [allProvinces[i], allProvinces[j]] = [allProvinces[j], allProvinces[i]];
    }
    
    // Fallback if not enough provinces in data
    while(allProvinces.length < 54 && allProvinces.length > 0) {
        allProvinces = allProvinces.concat(allProvinces);
    }
    const selectedProvs = allProvinces.slice(0, 54);
    
    const board = {};
    board[1] = { type: 'province', name: selectedProvs[0], icon: '🏁' };
    board[80] = { type: 'province', name: selectedProvs[53], icon: '🏆' };
    
    let availableSpaces = [];
    for (let i = 2; i <= 79; i++) availableSpaces.push(i);
    for (let i = availableSpaces.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [availableSpaces[i], availableSpaces[j]] = [availableSpaces[j], availableSpaces[i]];
    }

    // Allocate Gas Stations (10)
    for(let i=0; i<10; i++) {
        board[availableSpaces.pop()] = { type: 'gas', name: 'ปั๊มน้ำมัน', icon: '⛽' };
    }
    
    // Allocate Rest Stops (16 total, 2 of each)
    REST_TYPES.forEach(rest => {
        for(let i=0; i<2; i++) {
            board[availableSpaces.pop()] = { type: 'rest', name: rest.name, icon: rest.icon };
        }
    });

    // Allocate remaining 52 provinces
    let provIndex = 1;
    while(availableSpaces.length > 0) {
        const sp = availableSpaces.pop();
        board[sp] = { type: 'province', name: selectedProvs[provIndex], icon: '🏙️' };
        provIndex++;
    }

    return board;
}

window.createRallyRoom = function() {
    if (!myPlayerName) myPlayerName = 'นักแข่ง 1';
    playAudio('1.mp3');
    isStartingGame = false;

    const counterRef = ref(db, 'games/RallyThai/room_counter');
    runTransaction(counterRef, (cur) => (cur || 0) + 1).then((res) => {
        if (res.committed) {
            const count = res.snapshot.val();
            currentRoomId = 'Rally' + String(count).padStart(5, '0');
            myPlayerId = 'p1';
            isHost = true;
            const roomRef = ref(db, `games/RallyThai/rooms/${currentRoomId}`);
            onDisconnect(roomRef).remove();
            
            set(roomRef, {
                status: 'waiting',
                players: {
                    p1: { name: myPlayerName, avatar: CAR_POOL[0], isBot: false, pos: 1, fuel: 12 }
                },
                botCount: 0,
                turnIndex: 0,
                boardConfig: null,
                lastAction: { msg: '', ts: 0 },
                questionState: null
            });

            document.getElementById('lobby-menu-section').style.display = 'none';
            document.getElementById('lobby-room-section').style.display = 'block';
            setupRoomListener();
            announceSR(`สร้างห้องสำเร็จ รหัสห้องคือ ${currentRoomId}`);
        }
    });
};

window.joinRallyRoom = function(rId) {
    if (!myPlayerName) myPlayerName = 'นักแข่งใหม่';
    currentRoomId = rId;
    isHost = false;

    const roomRef = ref(db, `games/RallyThai/rooms/${currentRoomId}`);
    runTransaction(roomRef, (roomData) => {
        if (!roomData || roomData.status !== 'waiting') return;
        const pKeys = Object.keys(roomData.players || {});
        if (pKeys.length >= 6) return;
        let nextSlot = 'p1';
        for (let i = 1; i <= 6; i++) {
            if (!pKeys.includes('p' + i)) { nextSlot = 'p' + i; break; }
        }
        const usedAvatars = Object.values(roomData.players).map(p => p.avatar.icon);
        const availAvatars = CAR_POOL.filter(a => !usedAvatars.includes(a.icon));
        
        if (!roomData.players) roomData.players = {};
        roomData.players[nextSlot] = {
            name: myPlayerName, avatar: availAvatars[0] || CAR_POOL[0], isBot: false, pos: 1, fuel: 12
        };
        myPlayerId = nextSlot;
        roomData.lastAction = {
            msg: `${myPlayerName} เข้าร่วมห้องสำเร็จ`,
            audioKeys: ['select.mp3'],
            ts: Date.now() + Math.random()
        };
        return roomData;
    }).then((res) => {
        if (res.committed) {
            onDisconnect(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${myPlayerId}`)).update({
                isBot: true, name: "บอท" + myPlayerName
            });
            document.getElementById('lobby-menu-section').style.display = 'none';
            document.getElementById('lobby-room-section').style.display = 'block';
            setupRoomListener();
        }
    });
};

window.adjustBot = function(delta) {
    if (!isHost || !gameState) return;
    playAudio('select.mp3');
    const curBots = gameState.botCount || 0;
    const humanCount = Object.values(gameState.players).filter(p => !p.isBot).length;
    const totalCount = Object.keys(gameState.players || {}).length;
    const newBots = curBots + delta;
    
    if (delta > 0 && totalCount >= 6) return;
    if (delta < 0 && curBots <= 0) return;

    const roomRef = ref(db, `games/RallyThai/rooms/${currentRoomId}`);
    const updatedPlayers = { ...gameState.players };
    let actionMsg = '';

    if (delta > 0) {
        let botSlot = 'p1';
        for (let i = 1; i <= 6; i++) {
            if (!updatedPlayers['p' + i]) { botSlot = 'p' + i; break; }
        }
        const botName = `บอทนักซิ่ง ${newBots}`;
        const usedAvatars = Object.values(updatedPlayers).map(p => p.avatar.icon);
        const availAvatars = CAR_POOL.filter(a => !usedAvatars.includes(a.icon));
        updatedPlayers[botSlot] = {
            name: botName, avatar: availAvatars[0] || CAR_POOL[1], isBot: true, pos: 1, fuel: 12
        };
        actionMsg = `เพิ่ม ${botName} เข้าสู่ห้องแข่งแล้ว`;
    } else {
        const botSlots = Object.keys(updatedPlayers).filter(k => updatedPlayers[k].isBot);
        if (botSlots.length > 0) {
            const removedSlot = botSlots[botSlots.length - 1];
            const removedBotName = updatedPlayers[removedSlot].name;
            delete updatedPlayers[removedSlot];
            actionMsg = `นำ ${removedBotName} ออกจากห้องแข่งแล้ว`;
        }
    }

    const ts = Date.now() + Math.random();
    update(roomRef, { 
        botCount: newBots, 
        players: updatedPlayers,
        lastAction: { msg: actionMsg, ts: ts }
    });
};

function setupRoomListener() {
    const roomRef = ref(db, `games/RallyThai/rooms/${currentRoomId}`);
    roomListener = onValue(roomRef, (snapshot) => {
        gameState = snapshot.val();
        if (!gameState) {
            announceSR("เจ้าของห้องปิดเกม", 'assertive');
            setTimeout(() => { window.location.reload(); }, 2000);
            return;
        }

        updateLobbyUI();

        if (gameState.lastAction && gameState.lastAction.ts > localLastActionTs) {
            localLastActionTs = gameState.lastAction.ts;
            let actionToQueue = { ...gameState.lastAction };
            if (actionToQueue.wrongPId && actionToQueue.wrongPId === myPlayerId && myLastCorrectAnswerText) {
                actionToQueue.msg = actionToQueue.msg.replace('ตอบผิด!', `ตอบผิด! ${myLastCorrectAnswerText} เป็นคำตอบที่ถูกต้องนะ`);
            }
            actionQueue.push(actionToQueue);
            processActionQueue();
        }

        if (gameState.status === 'playing') {
            if (document.getElementById('screen-game').classList.contains('active-screen') === false) {
                showStartGameAnimation();
            } else {
                updateGameUI();
                checkQuestionState();
            }
        } else if (gameState.status === 'ended') {
            if (!isShowingWinnerScene) showWinnerAnimationAndResult();
        }
    });
}

function updateLobbyUI() {
    if (!gameState) return;
    document.getElementById('lobby-room-title').textContent = `ห้องแข่ง: ${currentRoomId}`;
    const pList = Object.values(gameState.players || {});
    document.getElementById('participant-count').textContent = `${pList.length}/6`;
    
    const listContainer = document.getElementById('participant-list');
    listContainer.innerHTML = '';
    pList.forEach(p => {
        const card = document.createElement('div');
        card.className = 'participant-card';
        card.innerHTML = `<span class="participant-avatar">${p.avatar.icon}</span> <span>${p.name} ${p.isBot ? '(บอท)' : ''}</span>`;
        listContainer.appendChild(card);
    });

    if (isHost) {
        document.getElementById('host-controls').style.display = 'block';
        const startBtn = document.getElementById('btn-start-game');
        startBtn.style.display = 'block';
        startBtn.disabled = isStartingGame || gameState.status !== 'waiting' || pList.length < 2;
        if (isStartingGame) {
            startBtn.classList.add('dim');
            startBtn.style.opacity = '0.5';
        }
        
        const botCountDisplay = document.getElementById('bot-count-display');
        if (botCountDisplay) botCountDisplay.textContent = gameState.botCount || 0;

        const addBotBtn = document.getElementById('btn-add-bot');
        if (addBotBtn) addBotBtn.disabled = pList.length >= 6;

        const removeBotBtn = document.getElementById('btn-remove-bot');
        if (removeBotBtn) removeBotBtn.disabled = (gameState.botCount || 0) <= 0;
    }
}

window.startRallyGame = function() {
    if (!isHost || isStartingGame) return;
    isStartingGame = true;
    const startBtn = document.getElementById('btn-start-game');
    if (startBtn) {
        startBtn.disabled = true;
        startBtn.classList.add('dim');
        startBtn.style.opacity = '0.5';
    }
    try {
        const boardConfig = generateRallyBoard();
        const pKeys = Object.keys(gameState.players || {});
        for (let i = pKeys.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [pKeys[i], pKeys[j]] = [pKeys[j], pKeys[i]];
        }
        update(ref(db, `games/RallyThai/rooms/${currentRoomId}`), {
            status: 'playing', boardConfig, turnIndex: 0, playerOrder: pKeys
        });
    } catch (err) {
        isStartingGame = false;
        if (startBtn) {
            startBtn.disabled = false;
            startBtn.classList.remove('dim');
            startBtn.style.opacity = '1';
        }
    }
};

async function showStartGameAnimation() {
    stopBGM();
    const overlay = document.getElementById('anim-start-overlay');
    overlay.style.display = 'flex';
    announceSR('การแข่งขันเริ่มขึ้นแล้ว!', 'assertive');
    
    await playAudio('start.mp3');
    playAudio('bgm.mp3', true);
    
    overlay.style.display = 'none';
    switchScreen('screen-game', 'game-status-bar');
    renderBoardGrid();
    updateGameUI();
}

function renderBoardGrid() {
    const grid = document.getElementById('rally-board');
    grid.innerHTML = '';
    grid.setAttribute('aria-hidden', 'true'); // ซ่อนกระดานทั้งหมดจาก Screen Reader เพื่อไม่ให้อ่านทีละช่อง

    const config = gameState.boardConfig || {};

    for (let row = 0; row < 8; row++) {
        for (let col = 0; col < 10; col++) {
            let num = (row % 2 === 0) ? (row * 10 + col + 1) : (row * 10 + (10 - col));
            const cell = document.createElement('div');
            cell.className = 'cell';
            cell.id = `cell-${num}`;
            
            let sp = config[num];
            if(sp) {
                if(sp.type === 'gas') cell.classList.add('cell-gas');
                else if(sp.type === 'rest') cell.classList.add('cell-rest');
                else cell.classList.add('cell-province');
                if(num === 1) cell.classList.add('cell-start');
                if(num === 80) cell.classList.add('cell-finish');
                
                cell.innerHTML = `<span class="cell-num">${num}</span>
                                  <span class="cell-icon">${sp.icon}</span>
                                  <div class="cell-name">${sp.name}</div>
                                  <div class="pawns-holder" id="pawns-holder-${num}"></div>`;
            }
            grid.appendChild(cell);
        }
    }
}

function updateGameUI() {
    if (!gameState || gameState.status !== 'playing') return;
    const pOrder = gameState.playerOrder || Object.keys(gameState.players);
    const pId = pOrder[gameState.turnIndex];
    const currentP = gameState.players[pId];

    document.getElementById('game-status-bar').textContent = `เทิร์นของ: ${currentP.avatar.icon} ${currentP.name}`;

    const rollBtn = document.getElementById('btn-roll-dice');
    if (pId === myPlayerId && !currentP.isBot && !gameState.turnExecuting && !gameState.questionState) {
        rollBtn.disabled = false;
        if (!myTurnTimer) myTurnTimer = setTimeout(() => { if (!rollBtn.disabled) window.handleRollDice(true); }, 40000);
    } else {
        rollBtn.disabled = true;
        if (myTurnTimer) { clearTimeout(myTurnTimer); myTurnTimer = null; }
    }

    if (lastAnnouncedTurnKey !== `${gameState.turnIndex}_${pId}`) {
        lastAnnouncedTurnKey = `${gameState.turnIndex}_${pId}`;
        
        // ประกาศเสียง SR ทันทีโดยไม่ต้องรอเสียงประกอบเพื่อความฉับไว
        announceSR(`ถึงเทิร์นของ ${currentP.name}`);

        // ดำเนินการเล่นเสียงประกอบคู่ขนานไปพร้อมกัน
        (async () => {
            await playAudio('turn.mp3');
            if (pId === myPlayerId) {
                playAudio('abc.mp3');
            }

            // จัดการดึงโฟกัสหลังจากเสียงจบอย่างเหมาะสม
            if (pId === myPlayerId && !currentP.isBot) {
                const btn = document.getElementById('btn-roll-dice');
                if (btn && !btn.disabled) btn.focus();
            }
        })();
    }

    // จัดกลุ่มสถานะสำหรับ Screen Reader ตามแนวทางที่กระชับ ไม่ให้รบกวนการแสดงผลจริง
    const cardsContainer = document.getElementById('player-status-cards');
    cardsContainer.innerHTML = '';
    cardsContainer.setAttribute('role', 'region');
    cardsContainer.setAttribute('aria-label', 'สถานะนักแข่งทั้งหมด');
    
    let activePlayerSummary = '';
    let otherPlayersSummary = [];

    pOrder.forEach((k) => {
        const p = gameState.players[k];
        const card = document.createElement('div');
        card.className = `p-status-card ${k === pId ? 'active-turn' : ''}`;
        card.setAttribute('aria-hidden', 'true'); // ซ่อนการ์ดแต่ละใบจาก SR ไม่ให้อ่านแยกกระจัดกระจาย
        card.innerHTML = `<div><strong>${p.avatar.icon} ${p.name}</strong></div><div>ช่อง ${p.pos} | น้ำมัน ${p.fuel}/12</div>`;
        cardsContainer.appendChild(card);
        
        const summaryItem = `${p.name} ช่อง ${p.pos} น้ำมัน ${p.fuel} ขีด`;
        if (k === myPlayerId) {
            activePlayerSummary = summaryItem;
        } else {
            otherPlayersSummary.push(summaryItem);
        }
    });

    let srTextParts = [];
    if (activePlayerSummary) srTextParts.push(activePlayerSummary);
    if (otherPlayersSummary.length > 0) srTextParts.push(otherPlayersSummary.join(' '));
    
    const srGroupElement = document.createElement('div');
    srGroupElement.className = 'sr-only';
    srGroupElement.textContent = srTextParts.join(' ');
    cardsContainer.appendChild(srGroupElement);

    document.querySelectorAll('.pawns-holder').forEach(h => h.innerHTML = '');
    pOrder.forEach((k) => {
        const p = gameState.players[k];
        const holder = document.getElementById(`pawns-holder-${p.pos}`);
        if (holder) {
            const token = document.createElement('div');
            token.className = 'pawn-token';
            token.textContent = p.avatar.icon;
            holder.appendChild(token);
        }
    });

    if (isHost && currentP.isBot && !gameState.turnExecuting && !gameState.questionState) {
        update(ref(db), { [`games/RallyThai/rooms/${currentRoomId}/turnExecuting`]: true });
        setTimeout(() => { executeTurnAsync(pId); }, 1500);
    }
}

async function syncActionEmit(msg, audioKeys = [], wrongPId = null) {
    const ts = Date.now() + Math.random();
    const actionObj = { msg, ts, audioKeys };
    if (wrongPId) actionObj.wrongPId = wrongPId;
    await update(ref(db), { [`games/RallyThai/rooms/${currentRoomId}/lastAction`]: actionObj });
    await delayAsync(600);
}

window.handleRollDice = function(isAuto = false) {
    document.getElementById('btn-roll-dice').disabled = true;
    update(ref(db), { [`games/RallyThai/rooms/${currentRoomId}/turnExecuting`]: true });
    executeTurnAsync(myPlayerId, isAuto);
};

async function executeTurnAsync(pId, isAuto = false) {
    const pData = gameState.players[pId];
    
    // Check Fuel Rule - น้ำมันหมดเล่น box3.mp3
    if (pData.fuel < 3) {
        const newFuel = Math.min(12, pData.fuel + 3);
        await syncActionEmit(`น้ำมันหมด! ${pData.name} ไม่สามารถเดินทางได้ ต้องหยุดพัก 1 เทิร์น และได้รับน้ำมัน 3 ขีด`, ['box3.mp3']);
        await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { fuel: newFuel });
        endTurn();
        return;
    }

    const newFuel = pData.fuel - 3;
    await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { fuel: newFuel });
    
    const diceRoll = Math.floor(Math.random() * 6) + 1;
    
    let currentPos = pData.pos;
    let targetPos = currentPos + diceRoll;
    if (targetPos > 80) targetPos = 80 - (targetPos - 80); // Bounce Rule
    
    const targetSp = gameState.boardConfig[targetPos];
    const cellName = targetSp ? targetSp.name : 'เส้นชัย';

    // ประกาศการทอยลูกเต๋าและการเดินรถพร้อมชื่อสถานที่เพื่อความกระชับ
    await syncActionEmit(`${pData.name} ใช้น้ำมัน 3 ขีด ทอยลูกเต๋าได้ ${diceRoll} เคลื่อนรถไปช่องที่ ${targetPos} ${cellName}`, ['dice.mp3', 'walk.mp3']);
    
    currentPos = targetPos;
    await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { pos: currentPos });

    if (currentPos === 80) {
        await handleWinGame(pId);
        return;
    }

    const sp = gameState.boardConfig[currentPos];
    
    if (sp.type === 'gas') {
        // เล่น box2.mp3 เฉพาะกรณีที่การเติมน้ำมันทำให้เชื้อเพลิงเต็มถังจริง (newFuel < 12)
        let gasAudio = newFuel < 12 ? ['box2.mp3'] : [];
        await syncActionEmit(`${pData.name} เข้าปั๊มน้ำมัน เติมน้ำมันฟรีเต็มถัง!`, gasAudio);
        await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { fuel: 12 });
        endTurn();
    } else if (sp.type === 'rest') {
        const fuelGain = Math.min(12, newFuel + 1);
        await syncActionEmit(`${pData.name} แวะพักที่ ${sp.name} ได้รับน้ำมัน 1 ขีด`, ['sabuy.mp3']);
        await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { fuel: fuelGain });
        endTurn();
    } else if (sp.type === 'province') {
        const qList = window.rallyQuestionsData[sp.name];
        if (qList && qList.length > 0) {
            // แจ้งให้ทุกคนทราบเพียงว่ากำลังตอบคำถาม (โดยไม่เห็นโจทย์) เดินถึงจังหวัดปกติเล่น box1.mp3
            await syncActionEmit(`${pData.name} กำลังตอบคำถามจากจังหวัด ${sp.name}`, ['box1.mp3']);

await delayAsync(1400);
            const q = qList[Math.floor(Math.random() * qList.length)];
            
            // Shuffle options
            let opts = [...q.options];
            for (let i = opts.length - 1; i > 0; i--) {
                const j = Math.floor(Math.random() * (i + 1));
                [opts[i], opts[j]] = [opts[j], opts[i]];
            }
            
            await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/questionState`), {
                pId: pId, prov: sp.name, question: q.question, options: opts, correctId: q.correctOptionId,
                answeredBy: null, selectedOpt: null, processing: false, isAuto: !!isAuto
            });
            
            if(pData.isBot && isHost) {
                // Bot logic: pick random answer
                setTimeout(() => {
                    const pickedOpt = opts[Math.floor(Math.random() * opts.length)];
                    handleAnswer(pId, pickedOpt.id);
                }, 4000);
            }
        } else {
            // เดินถึงจังหวัดปกติไม่มีคำถามเล่น box1.mp3
            await syncActionEmit(`ถึง ${sp.name} (ไม่มีคำถามในฐานข้อมูล) แวะพักผ่อนเฉยๆ`, ['box1.mp3']);
            endTurn();
        }
    }
}

function checkQuestionState() {
    const qState = gameState.questionState;
    const modal = document.getElementById('question-modal-overlay');

    if (isHost && qState && qState.answeredBy && !qState.processing) {
        update(ref(db, `games/RallyThai/rooms/${currentRoomId}/questionState`), { processing: true });
        processAnswer(qState.answeredBy, qState.selectedOpt);
        return;
    }
    
    // ผู้เล่นที่เป็นเจ้าของเทิร์นเท่านั้นจะเห็นคำถาม
    if (qState && qState.pId === myPlayerId && !gameState.players[myPlayerId].isBot) {
        const correctOpt = qState.options ? qState.options.find(o => o.id === qState.correctId) : null;
        if (correctOpt) {
            myLastCorrectAnswerText = correctOpt.text;
        }
        if (modal.style.display !== 'flex') {
            document.getElementById('question-province').textContent = `ขับรถมาถึงจังหวัด ${qState.prov}`;
            document.getElementById('question-text').textContent = qState.question;
            const optsContainer = document.getElementById('question-options');
            optsContainer.innerHTML = '';
            
            qState.options.forEach(opt => {
                const btn = document.createElement('button');
                btn.textContent = opt.text;
                btn.onclick = () => {
                    if (autoAnswerTimer) { clearTimeout(autoAnswerTimer); autoAnswerTimer = null; }
                    modal.style.display = 'none';
                    // เมื่อตอบเสร็จ ให้กลับโฟกัสเข้าสู่หน้าหลักของเกมเพื่อเล่นต่อ
                    const gameHeader = document.getElementById('game-status-bar');
                    if (gameHeader) {
                        gameHeader.setAttribute('tabindex', '-1');
                        gameHeader.focus();
                    }
                    handleAnswer(myPlayerId, opt.id);
                };
                optsContainer.appendChild(btn);
            });
            
            modal.style.display = 'flex';
            announceSR(`คำถามจังหวัด ${qState.prov}`);

            // ย้ายโฟกัสทันทีเมื่อเปิด Overlay คำถามและล็อคเป้าไว้ที่ Heading
            const heading = document.getElementById('question-province');
            if (heading) {
                heading.setAttribute('tabindex', '-1');
                heading.focus();
            }
        }

        if (qState.isAuto && !qState.answeredBy && !qState.processing && !autoAnswerTimer) {
            autoAnswerTimer = setTimeout(() => {
                autoAnswerTimer = null;
                if (gameState && gameState.questionState && gameState.questionState.pId === myPlayerId && !gameState.questionState.answeredBy) {
                    modal.style.display = 'none';
                    const gameHeader = document.getElementById('game-status-bar');
                    if (gameHeader) {
                        gameHeader.setAttribute('tabindex', '-1');
                        gameHeader.focus();
                    }
                    const opts = gameState.questionState.options;
                    if (opts && opts.length > 0) {
                        const randomOpt = opts[Math.floor(Math.random() * opts.length)];
                        handleAnswer(myPlayerId, randomOpt.id);
                    }
                }
            }, 1000);
        }
    } else {
        if (autoAnswerTimer) { clearTimeout(autoAnswerTimer); autoAnswerTimer = null; }
        modal.style.display = 'none';
    }
}

window.handleAnswer = async function(pId, selectedOptId) {
    if (autoAnswerTimer) { clearTimeout(autoAnswerTimer); autoAnswerTimer = null; }
    if (!gameState || !gameState.questionState || gameState.questionState.answeredBy) return;

    // Player answers and sends to Host via DB
    await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/questionState`), {
        answeredBy: pId, selectedOpt: selectedOptId
    });
    
    // If we are Host, process standard flow
    if (isHost && gameState && gameState.questionState && !gameState.questionState.processing) {
        update(ref(db, `games/RallyThai/rooms/${currentRoomId}/questionState`), { processing: true });
        processAnswer(pId, selectedOptId);
    }
};

async function processAnswer(pId, selectedOptId) {
    const qState = gameState.questionState;
    if (!qState) return;
    const pData = gameState.players[pId];
    const isCorrect = (selectedOptId === qState.correctId);
    
    // Clear question state so UI closes for everyone
    await update(ref(db, `games/RallyThai/rooms/${currentRoomId}`), { questionState: null });

    if (isCorrect) {
        const fuelGain = Math.min(12, pData.fuel + 2);
        const forwardMove = Math.floor(Math.random() * 4) + 2; // 2-5 spaces
        let targetPos = pData.pos + forwardMove;
        if (targetPos > 80) targetPos = 80 - (targetPos - 80);

        const targetSp = gameState.boardConfig[targetPos];
        const cellName = targetSp ? targetSp.name : 'เส้นชัย';

        // ตอบถูก เล่น shield.mp3 ตามด้วย walk.mp3
        await syncActionEmit(`${pData.name} ตอบถูก! ได้น้ำมัน 2 ขีด และเดินหน้า ${forwardMove} ช่อง ไปยังช่องที่ ${targetPos} ${cellName}`, ['shield.mp3', 'walk.mp3']); 
        await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { fuel: fuelGain });
        
        await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { pos: targetPos });
        
        if (targetPos === 80) { await handleWinGame(pId); return; }
        
    } else {
        const backMove = Math.floor(Math.random() * 3) + 1; // 1-3 spaces
        let targetPos = Math.max(1, pData.pos - backMove);
        
        const targetSp = gameState.boardConfig[targetPos];
        const cellName = targetSp ? targetSp.name : 'จุดเริ่มต้น';

        // ตอบผิด เล่น shieldno.mp3 ตามด้วย walk1.mp3
        await syncActionEmit(`${pData.name} ตอบผิด! ถอยหลัง ${backMove} ช่อง ไปยังช่องที่ ${targetPos} ${cellName}`, ['shieldno.mp3', 'walk1.mp3'], pId);
        
        await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { pos: targetPos });
    }
    
    endTurn();
}

async function endTurn() {
    // ป้องกันการอัปเดตเทิร์นจนกว่าคิวเหตุการณ์ภายในเครื่อง (ทั้งเสียงและ SR) จะทำงานเสร็จ ป้องกัน Race Condition
    while(isProcessingAction || actionQueue.length > 0) {
        await delayAsync(200);
    }
    
    const pKeys = gameState.playerOrder || Object.keys(gameState.players);
    const nextTurn = (gameState.turnIndex + 1) % pKeys.length;
    await delayAsync(1800); 
    await update(ref(db), { 
        [`games/RallyThai/rooms/${currentRoomId}/turnIndex`]: nextTurn,
        [`games/RallyThai/rooms/${currentRoomId}/turnExecuting`]: false
    });
}

async function handleWinGame(pId) {
    await update(ref(db), { 
        [`games/RallyThai/rooms/${currentRoomId}/status`]: 'ended',
        [`games/RallyThai/rooms/${currentRoomId}/winnerId`]: pId,
        [`games/RallyThai/rooms/${currentRoomId}/turnExecuting`]: false,
        [`games/RallyThai/rooms/${currentRoomId}/questionState`]: null
    });
}

async function showWinnerAnimationAndResult() {
    isShowingWinnerScene = true;
    stopBGM();
    const overlay = document.getElementById('anim-winner-overlay');
    const winner = gameState.players[gameState.winnerId];
    document.getElementById('winner-anim-character').textContent = winner.avatar.icon;
    document.getElementById('winner-anim-name').textContent = winner.name;
    overlay.style.display = 'flex';
    
    await playAudio('win.mp3');
    
    overlay.style.display = 'none';
    switchScreen('screen-result', 'result-title');
    document.getElementById('result-winner-text').textContent = `🎉 ${winner.avatar.icon} ${winner.name} ถึงเส้นชัยเป็นคนแรก!`;
    announceSR(`การแข่งขันสิ้นสุด ผู้ชนะคือ ${winner.name}`, 'assertive');
    
    const rankingsBox = document.getElementById('result-rankings');
    rankingsBox.innerHTML = '<h3>อันดับนักแข่ง:</h3>';
    const sorted = Object.values(gameState.players).sort((a, b) => b.pos - a.pos);
    sorted.forEach(p => { rankingsBox.innerHTML += `<p>${p.avatar.icon} ${p.name} ช่อง ${p.pos}</p>`; });
}

window.leaveRoom = function() {
    if (autoAnswerTimer) { clearTimeout(autoAnswerTimer); autoAnswerTimer = null; }
    if (currentRoomId && myPlayerId) {
        if (isHost) remove(ref(db, `games/RallyThai/rooms/${currentRoomId}`));
        else update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${myPlayerId}`), { isBot: true, name: "บอท" + myPlayerName });
    }
    stopBGM();
    myPlayerId = null; currentRoomId = null; isHost = false; gameState = null;
    isShowingWinnerScene = false; isStartingGame = false;
    myLastCorrectAnswerText = '';
    if (myTurnTimer) { clearTimeout(myTurnTimer); myTurnTimer = null; }
    document.getElementById('lobby-room-section').style.display = 'none';
    document.getElementById('lobby-menu-section').style.display = 'block';
};

window.returnToLobbyOrMain = function() {
    window.leaveRoom();
    switchScreen('screen-lobby', 'lobby-heading');
};

document.addEventListener('DOMContentLoaded', () => {
    initRoomListListener();
    const nameInput = document.getElementById('player-name-input');
    const confirmBtn = document.getElementById('btn-confirm-name');
    if (nameInput && confirmBtn) {
        nameInput.addEventListener('input', () => confirmBtn.disabled = nameInput.value.trim().length === 0);
        nameInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && nameInput.value.trim().length > 0) window.confirmNameAndEnterLobby();
        });
    }
});

// แป้นพิมพ์ลัดสำหรับผู้ใช้ Windows และ Screen Reader (Alt+R และ Alt+A)
document.addEventListener('keydown', (e) => {
    if (e.repeat) return;

    const targetTag = e.target ? e.target.tagName : '';
    if (targetTag === 'INPUT' || targetTag === 'TEXTAREA' || (e.target && e.target.isContentEditable)) {
        return;
    }

    if (!e.altKey || e.ctrlKey || e.metaKey) return;

    const key = e.key ? e.key.toLowerCase() : '';
    const code = e.code || '';

    // Alt + R: เรียกใช้งานปุ่มทอยลูกเต๋า (btn-roll-dice)
    if (key === 'r' || key === '®' || code === 'KeyR') {
        if (gameState && gameState.status === 'playing' && gameState.players) {
            const pOrder = gameState.playerOrder || Object.keys(gameState.players);
            const pId = pOrder[gameState.turnIndex];
            const currentP = gameState.players[pId];
            const rollBtn = document.getElementById('btn-roll-dice');

            if (pId === myPlayerId && currentP && !currentP.isBot && rollBtn && !rollBtn.disabled && !gameState.turnExecuting && !gameState.questionState) {
                e.preventDefault();
                rollBtn.click();
            }
        }
    }

    // Alt + A: ประกาศสถานะผู้เล่นทั้งหมดให้ Screen Reader ทราบ
    if (key === 'a' || key === 'å' || code === 'KeyA') {
        if (gameState && gameState.status === 'playing' && gameState.players) {
            e.preventDefault();
            const pOrder = gameState.playerOrder || Object.keys(gameState.players);
            let activePlayerSummary = '';
            let otherPlayersSummary = [];

            pOrder.forEach((k) => {
                const p = gameState.players[k];
                if (p) {
                    const summaryItem = `${p.name} ช่อง ${p.pos} น้ำมัน ${p.fuel} ขีด`;
                    if (k === myPlayerId) {
                        activePlayerSummary = summaryItem;
                    } else {
                        otherPlayersSummary.push(summaryItem);
                    }
                }
            });

            let srTextParts = [];
            if (activePlayerSummary) srTextParts.push(activePlayerSummary);
            if (otherPlayersSummary.length > 0) srTextParts.push(otherPlayersSummary.join(' '));

            const srText = srTextParts.join(' ');
            if (srText) {
                announceSR(srText);
            }
        }
    }
});