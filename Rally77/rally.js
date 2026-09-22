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
let isStartingGame = false;
let isShowingWinnerScene = false;

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
                    p1: { name: myPlayerName, avatar: CAR_POOL[0], isBot: false, pos: 1, fuel: 10 }
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
            name: myPlayerName, avatar: availAvatars[0] || CAR_POOL[0], isBot: false, pos: 1, fuel: 10
        };
        myPlayerId = nextSlot;
        return roomData;
    }).then((res) => {
        if (res.committed) {
            playAudio('select.mp3');
            onDisconnect(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${myPlayerId}`)).update({
                isBot: true, name: "บอท" + myPlayerName
            });
            document.getElementById('lobby-menu-section').style.display = 'none';
            document.getElementById('lobby-room-section').style.display = 'block';
            setupRoomListener();
            announceSR(`เข้าร่วมห้อง ${currentRoomId} เรียบร้อย`);
        }
    });
};

window.adjustBot = function(delta) {
    if (!isHost || !gameState) return;
    const curBots = gameState.botCount || 0;
    const humanCount = Object.values(gameState.players).filter(p => !p.isBot).length;
    const newBots = curBots + delta;
    if (newBots < 0 || (humanCount + newBots) > 6) return;

    const roomRef = ref(db, `games/RallyThai/rooms/${currentRoomId}`);
    const updatedPlayers = { ...gameState.players };

    if (delta > 0) {
        const botSlot = 'p' + (humanCount + newBots);
        const usedAvatars = Object.values(updatedPlayers).map(p => p.avatar.icon);
        const availAvatars = CAR_POOL.filter(a => !usedAvatars.includes(a.icon));
        updatedPlayers[botSlot] = {
            name: `บอทนักซิ่ง ${newBots}`, avatar: availAvatars[0] || CAR_POOL[1], isBot: true, pos: 1, fuel: 10
        };
    } else {
        const botSlots = Object.keys(updatedPlayers).filter(k => updatedPlayers[k].isBot);
        if (botSlots.length > 0) delete updatedPlayers[botSlots[botSlots.length - 1]];
    }

    update(roomRef, { botCount: newBots, players: updatedPlayers });
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
            announceSR(gameState.lastAction.msg, 'polite');
            if (gameState.lastAction.audioKeys) playAudioSequence(gameState.lastAction.audioKeys);
            
            if (gameState.lastAction.msg.includes('ทอยลูกเต๋าได้')) {
                const match = gameState.lastAction.msg.match(/ทอยลูกเต๋าได้\s*(\d+)/);
                if (match) document.getElementById('dice-visual').textContent = match[1];
            }
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
        document.getElementById('bot-count-display').textContent = gameState.botCount || 0;
    }
}

window.startRallyGame = function() {
    if (!isHost || isStartingGame) return;
    isStartingGame = true;
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
    }
};

function showStartGameAnimation() {
    stopBGM();
    const overlay = document.getElementById('anim-start-overlay');
    overlay.style.display = 'flex';
    announceSR('การแข่งขันเริ่มขึ้นแล้ว!', 'assertive');
    playAudio('start.mp3').then(() => { setTimeout(() => { playAudio('bgm.mp3', true); }, 200); });
    setTimeout(() => {
        overlay.style.display = 'none';
        switchScreen('screen-game', 'game-status-bar');
        renderBoardGrid();
        updateGameUI();
    }, 4800);
}

function renderBoardGrid() {
    const grid = document.getElementById('rally-board');
    grid.innerHTML = '';
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
                cell.setAttribute('aria-label', `ช่อง ${num} ${sp.name}`);
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
        announceSR(`ถึงเทิร์นของ ${currentP.name}`);
    }

    const cardsContainer = document.getElementById('player-status-cards');
    cardsContainer.innerHTML = '';
    pOrder.forEach((k) => {
        const p = gameState.players[k];
        const card = document.createElement('div');
        card.className = `p-status-card ${k === pId ? 'active-turn' : ''}`;
        card.innerHTML = `<div><strong>${p.avatar.icon} ${p.name}</strong></div><div>ช่อง ${p.pos} | น้ำมัน ${p.fuel}/10</div>`;
        cardsContainer.appendChild(card);
    });

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

async function syncActionEmit(msg, audioKeys = []) {
    const ts = Date.now() + Math.random();
    await update(ref(db), { [`games/RallyThai/rooms/${currentRoomId}/lastAction`]: { msg, ts, audioKeys } });
    await delayAsync(2200);
}

window.handleRollDice = function(isAuto = false) {
    document.getElementById('btn-roll-dice').disabled = true;
    update(ref(db), { [`games/RallyThai/rooms/${currentRoomId}/turnExecuting`]: true });
    executeTurnAsync(myPlayerId);
};

async function executeTurnAsync(pId) {
    const pData = gameState.players[pId];
    
    // Check Fuel Rule
    if (pData.fuel < 3) {
        const newFuel = Math.min(10, pData.fuel + 3);
        await syncActionEmit(`น้ำมันหมด! ${pData.name} ไม่สามารถเดินทางได้ ต้องหยุดพัก 1 เทิร์น และได้รับน้ำมัน 3 ขีด`);
        await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { fuel: newFuel });
        endTurn();
        return;
    }

    const newFuel = pData.fuel - 3;
    await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { fuel: newFuel });
    
    const diceRoll = Math.floor(Math.random() * 6) + 1;
    await syncActionEmit(`${pData.name} ใช้น้ำมัน 3 ขีด ทอยลูกเต๋าได้ ${diceRoll}`, ['dice.mp3']);
    
    let currentPos = pData.pos;
    let targetPos = currentPos + diceRoll;
    if (targetPos > 80) targetPos = 80 - (targetPos - 80); // Bounce Rule
    
    await syncActionEmit(`${pData.name} ขับรถไปช่อง ${targetPos}`, ['walk.mp3']); // Using walk.mp3 as movement sound based on prompt
    
    currentPos = targetPos;
    await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { pos: currentPos });

    if (currentPos === 80) {
        await handleWinGame(pId);
        return;
    }

    const sp = gameState.boardConfig[currentPos];
    
    if (sp.type === 'gas') {
        await syncActionEmit(`${pData.name} เข้าปั๊มน้ำมัน เติมน้ำมันฟรีเต็มถัง 10 ขีด!`, ['sabuy.mp3']);
        await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { fuel: 10 });
        endTurn();
    } else if (sp.type === 'rest') {
        const fuelGain = Math.min(10, newFuel + 1);
        await syncActionEmit(`${pData.name} แวะพักที่ ${sp.name} ได้รับน้ำมัน 1 ขีด`, ['sabuy.mp3']);
        await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { fuel: fuelGain });
        endTurn();
    } else if (sp.type === 'province') {
        // Province Question Rule
        const qList = window.rallyQuestionsData[sp.name];
        if (qList && qList.length > 0) {
            const q = qList[Math.floor(Math.random() * qList.length)];
            
            // Shuffle options
            let opts = [...q.options];
            for (let i = opts.length - 1; i > 0; i--) {
                const j = Math.floor(Math.random() * (i + 1));
                [opts[i], opts[j]] = [opts[j], opts[i]];
            }
            
            await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/questionState`), {
                pId: pId, prov: sp.name, question: q.question, options: opts, correctId: q.correctOptionId
            });
            
            if(pData.isBot && isHost) {
                // Bot logic: pick random answer
                setTimeout(() => {
                    const pickedOpt = opts[Math.floor(Math.random() * opts.length)];
                    handleAnswer(pId, pickedOpt.id);
                }, 4000);
            }
        } else {
            await syncActionEmit(`ถึง ${sp.name} (ไม่มีคำถามในฐานข้อมูล) แวะพักผ่อนเฉยๆ`);
            endTurn();
        }
    }
}

function checkQuestionState() {
    const qState = gameState.questionState;
    const modal = document.getElementById('question-modal-overlay');
    
    if (qState && qState.pId === myPlayerId && !gameState.players[myPlayerId].isBot) {
        if (modal.style.display !== 'flex') {
            document.getElementById('question-province').textContent = `คำถามจังหวัด ${qState.prov}`;
            document.getElementById('question-text').textContent = qState.question;
            const optsContainer = document.getElementById('question-options');
            optsContainer.innerHTML = '';
            
            qState.options.forEach(opt => {
                const btn = document.createElement('button');
                btn.textContent = opt.text;
                btn.onclick = () => {
                    modal.style.display = 'none';
                    handleAnswer(myPlayerId, opt.id);
                };
                optsContainer.appendChild(btn);
            });
            
            modal.style.display = 'flex';
            announceSR(`มีคำถามจากจังหวัด ${qState.prov}: ${qState.question}`);
        }
    } else {
        modal.style.display = 'none';
    }
}

window.handleAnswer = async function(pId, selectedOptId) {
    // Player answers and sends to Host via DB
    await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/questionState`), {
        answeredBy: pId, selectedOpt: selectedOptId
    });
    
    // If we are Host, we process the answer
    if (isHost) processAnswer(pId, selectedOptId);
};

async function processAnswer(pId, selectedOptId) {
    const qState = gameState.questionState;
    const pData = gameState.players[pId];
    const isCorrect = (selectedOptId === qState.correctId);
    
    // Clear question state so UI closes for everyone
    await update(ref(db, `games/RallyThai/rooms/${currentRoomId}`), { questionState: null });

    if (isCorrect) {
        const fuelGain = Math.min(10, pData.fuel + 2);
        const forwardMove = Math.floor(Math.random() * 4) + 2; // 2-5 spaces
        await syncActionEmit(`${pData.name} ตอบถูก! เติมน้ำมัน 2 ขีด และโชคดีได้ซิ่งเดินหน้า ${forwardMove} ช่อง`, ['win.mp3']); // Reusing win briefly or forward
        await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { fuel: fuelGain });
        
        let targetPos = pData.pos + forwardMove;
        if (targetPos > 80) targetPos = 80 - (targetPos - 80);
        await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { pos: targetPos });
        
        if (targetPos === 80) { await handleWinGame(pId); return; }
        
    } else {
        const backMove = Math.floor(Math.random() * 3) + 1; // 1-3 spaces
        await syncActionEmit(`${pData.name} ตอบผิด! ไม่ได้น้ำมันเพิ่ม และหลงทางถอยหลัง ${backMove} ช่อง`, ['shieldno.mp3']);
        
        let targetPos = Math.max(1, pData.pos - backMove);
        await update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${pId}`), { pos: targetPos });
    }
    
    endTurn();
}

async function endTurn() {
    const pKeys = gameState.playerOrder || Object.keys(gameState.players);
    const nextTurn = (gameState.turnIndex + 1) % pKeys.length;
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

function showWinnerAnimationAndResult() {
    isShowingWinnerScene = true;
    stopBGM();
    const overlay = document.getElementById('anim-winner-overlay');
    const winner = gameState.players[gameState.winnerId];
    document.getElementById('winner-anim-character').textContent = winner.avatar.icon;
    document.getElementById('winner-anim-name').textContent = winner.name;
    overlay.style.display = 'flex';
    setTimeout(() => { playAudio('win.mp3'); }, 100);
    
    setTimeout(() => {
        overlay.style.display = 'none';
        switchScreen('screen-result', 'result-title');
        document.getElementById('result-winner-text').textContent = `🎉 ${winner.avatar.icon} ${winner.name} ถึงเส้นชัยเป็นคนแรก!`;
        announceSR(`การแข่งขันสิ้นสุด ผู้ชนะคือ ${winner.name}`, 'assertive');
        
        const rankingsBox = document.getElementById('result-rankings');
        rankingsBox.innerHTML = '<h3>อันดับนักแข่ง:</h3>';
        const sorted = Object.values(gameState.players).sort((a, b) => b.pos - a.pos);
        sorted.forEach(p => { rankingsBox.innerHTML += `<p>${p.avatar.icon} ${p.name} ช่อง ${p.pos}</p>`; });
    }, 5000);
}

window.leaveRoom = function() {
    if (currentRoomId && myPlayerId) {
        if (isHost) remove(ref(db, `games/RallyThai/rooms/${currentRoomId}`));
        else update(ref(db, `games/RallyThai/rooms/${currentRoomId}/players/${myPlayerId}`), { isBot: true, name: "บอท" + myPlayerName });
    }
    stopBGM();
    myPlayerId = null; currentRoomId = null; isHost = false; gameState = null;
    isShowingWinnerScene = false; isStartingGame = false;
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