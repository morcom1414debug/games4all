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
let myPlayerId = null;
let currentRoomId = null;
let isHost = false;
let gameState = null;
let roomListener = null;
let roomListListener = null;
let speechQueue = [];
let isSpeaking = false;

// Audio Context Web Synthesizer
const audioCtx = new (window.AudioContext || window.webkitAudioContext)();

function playSynthSound(type) {
    if (audioCtx.state === 'suspended') audioCtx.resume();
    const osc = audioCtx.createOscillator();
    const gain = audioCtx.createGain();
    osc.connect(gain);
    gain.connect(audioCtx.destination);

    const now = audioCtx.currentTime;
    if (type === 'roll') {
        osc.frequency.setValueAtTime(300, now);
        osc.frequency.exponentialRampToValueAtTime(150, now + 0.15);
        gain.gain.setValueAtTime(0.3, now);
        gain.gain.linearRampToValueAtTime(0.01, now + 0.15);
        osc.start(now);
        osc.stop(now + 0.15);
    } else if (type === 'move') {
        osc.frequency.setValueAtTime(400, now);
        osc.frequency.exponentialRampToValueAtTime(600, now + 0.08);
        gain.gain.setValueAtTime(0.2, now);
        gain.gain.linearRampToValueAtTime(0.01, now + 0.08);
        osc.start(now);
        osc.stop(now + 0.08);
    } else if (type === 'treasure' || type === 'key') {
        osc.frequency.setValueAtTime(523.25, now);
        osc.frequency.setValueAtTime(659.25, now + 0.1);
        gain.gain.setValueAtTime(0.3, now);
        gain.gain.linearRampToValueAtTime(0.01, now + 0.25);
        osc.start(now);
        osc.stop(now + 0.25);
    } else if (type === 'bad_event') {
        osc.frequency.setValueAtTime(200, now);
        osc.frequency.linearRampToValueAtTime(100, now + 0.3);
        gain.gain.setValueAtTime(0.4, now);
        gain.gain.linearRampToValueAtTime(0.01, now + 0.3);
        osc.start(now);
        osc.stop(now + 0.3);
    } else if (type === 'win') {
        osc.frequency.setValueAtTime(523.25, now);
        osc.frequency.setValueAtTime(659.25, now + 0.15);
        osc.frequency.setValueAtTime(783.99, now + 0.3);
        gain.gain.setValueAtTime(0.4, now);
        gain.gain.linearRampToValueAtTime(0.01, now + 0.5);
        osc.start(now);
        osc.stop(now + 0.5);
    }
}

// ARIA Queue Announcement Manager
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
            setTimeout(() => {
                isSpeaking = false;
                processSpeechQueue();
            }, 800);
        }, 50);
    } else {
        isSpeaking = false;
    }
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
            }
        }
    }
}

// Initialize Room Listener for Main Menu
function initRoomListListener() {
    const roomsRef = ref(db, 'games/Adventure80/rooms');
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
                item.className = 'room-item';
                const pCount = rData.players ? Object.keys(rData.players).length : 0;
                item.innerHTML = `<span><strong>ห้อง ${rId}</strong> (${pCount}/6 คน)</span>`;
                item.onclick = () => window.joinAdventureRoom(rId);
                container.appendChild(item);
            }
        }
        if (count === 0) {
            container.innerHTML = '<div style="color:#a0a5b5; padding:10px;">ไม่มีห้องที่กำลังรอผู้เล่นอยู่...</div>';
        }
    });
}

// Board Random Generator (Host Authoritative)
function generateBoardConfiguration() {
    const totalSpaces = 80;
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
    const usedSpaces = new Set();

    // Select 6 Keys and 6 Doors with constraints (Door > 10, Door - Key >= 6)
    for (let k = 0; k < 6; k++) {
        let keySpace = available.pop();
        while (keySpace >= 70 && available.length > 0) {
            available.unshift(keySpace);
            keySpace = available.pop();
        }

        let possibleDoorSpaces = available.filter(s => s > 10 && (s - keySpace) >= 6);
        if (possibleDoorSpaces.length === 0) {
            keySpace = 12 + k * 2;
            possibleDoorSpaces = [keySpace + 7];
        }
        const doorSpace = possibleDoorSpaces[Math.floor(Math.random() * possibleDoorSpaces.length)];
        
        const idx = available.indexOf(doorSpace);
        if (idx !== -1) available.splice(idx, 1);

        usedSpaces.add(keySpace);
        usedSpaces.add(doorSpace);

        const dest = Math.min(79, doorSpace + Math.floor(Math.random() * 8) + 3);
        keys.push(keySpace);
        doors.push({ space: doorSpace, dest });
    }

    // Allocate remaining special spaces
    const layout = {};
    keys.forEach(k => layout[k] = { type: 'key' });
    doors.forEach(d => layout[d.space] = { type: 'door', dest: d.dest });

    const specialTypes = [
        { type: 'treasure', count: 8 },
        { type: 'forward', count: 6 },
        { type: 'trap', count: 7 },
        { type: 'water', count: 5 },
        { type: 'ghost', count: 5 },
        { type: 'warp', count: 4 }
    ];

    specialTypes.forEach(st => {
        for (let c = 0; c < st.count; c++) {
            if (available.length > 0) {
                const sp = available.pop();
                usedSpaces.add(sp);
                let effectVal = 0;
                if (st.type === 'forward') effectVal = Math.floor(Math.random() * 5) + 2; // +2..+6
                else if (st.type === 'trap') effectVal = -(Math.floor(Math.random() * 5) + 2); // -2..-6
                else if (st.type === 'water') effectVal = -(Math.floor(Math.random() * 5) + 1); // -1..-5
                else if (st.type === 'warp') effectVal = Math.random() > 0.5 ? (Math.floor(Math.random() * 11) + 10) : -(Math.floor(Math.random() * 11) + 10);
                
                layout[sp] = { type: st.type, val: effectVal };
            }
        }
    });

    return layout;
}

// Create Room Handler
window.createAdventureRoom = function() {
    const nameInput = document.getElementById('player-name-input');
    const pName = nameInput.value.trim() || 'ผู้เล่น 1';

    const counterRef = ref(db, 'games/Adventure80/room_counter');
    runTransaction(counterRef, (cur) => (cur || 0) + 1).then((res) => {
        if (res.committed) {
            const count = res.snapshot.val();
            currentRoomId = 'Adventure' + String(count).padStart(5, '0');
            myPlayerId = 'p1';
            isHost = true;

            const roomRef = ref(db, `games/Adventure80/rooms/${currentRoomId}`);
            onDisconnect(roomRef).remove();

            const initialPlayers = {
                p1: { name: pName, animal: ANIMAL_POOL[0], isBot: false, pos: 1, keys: 0, armor: 0 }
            };

            set(roomRef, {
                status: 'waiting',
                players: initialPlayers,
                botCount: 0,
                turnIndex: 0,
                boardConfig: null
            });

            setupRoomListener();
            switchScreen('screen-lobby', 'lobby-room-title');
            announceSR(`สร้างห้องสำเร็จ รหัสห้องคือ ${currentRoomId}`);
        }
    });
};

// Join Room Handler
window.joinAdventureRoom = function(rId) {
    const nameInput = document.getElementById('player-name-input');
    const pName = nameInput.value.trim() || 'ผู้เล่นเข้าใหม่';

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
            name: pName,
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
            onDisconnect(ref(db, `games/Adventure80/rooms/${currentRoomId}/players/${myPlayerId}`)).remove();
            setupRoomListener();
            switchScreen('screen-lobby', 'lobby-room-title');
            announceSR(`เข้าร่วมห้อง ${currentRoomId} เรียบร้อยแล้ว`);
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
            alert('ห้องเกมนี้ถูกปิดแล้ว');
            window.leaveRoom();
            return;
        }

        updateLobbyUI();

        if (gameState.status === 'playing') {
            if (document.getElementById('screen-game').classList.contains('active-screen') === false) {
                switchScreen('screen-game', 'game-status-bar');
                renderBoardGrid();
            }
            updateGameUI();
        } else if (gameState.status === 'ended') {
            showResultScreen();
        }
    });
}

// Update Lobby Interface
function updateLobbyUI() {
    if (!gameState) return;
    document.getElementById('lobby-room-title').textContent = `ห้องเกม: ${currentRoomId}`;
    
    const pList = Object.values(gameState.players || {});
    document.getElementById('participant-count').textContent = `${pList.length}/6`;

    const listContainer = document.getElementById('participant-list');
    listContainer.innerHTML = '';
    pList.forEach(p => {
        const card = document.createElement('div');
        card.className = 'participant-card';
        card.innerHTML = `<span class="participant-avatar">${p.animal.icon}</span> <span>${p.name} ${p.isBot ? '(บอท)' : ''}</span>`;
        listContainer.appendChild(card);
    });

    const hostPanel = document.getElementById('host-controls');
    const startBtn = document.getElementById('btn-start-game');

    if (isHost) {
        hostPanel.style.display = 'block';
        startBtn.style.display = 'block';
        startBtn.disabled = pList.length < 2;
        document.getElementById('bot-count-display').textContent = gameState.botCount || 0;
    } else {
        hostPanel.style.display = 'none';
        startBtn.style.display = 'none';
    }
}

// Start Game Handler (Host Only)
window.startAdventureGame = function() {
    if (!isHost) return;
    const boardConfig = generateBoardConfiguration();
    const roomRef = ref(db, `games/Adventure80/rooms/${currentRoomId}`);

    update(roomRef, {
        status: 'playing',
        boardConfig,
        turnIndex: 0
    });
};

// Render Serpentine Board Grid 8x10
function renderBoardGrid() {
    const grid = document.getElementById('adventure-board');
    grid.innerHTML = '';
    const config = gameState.boardConfig || {};

    // 8 rows, 10 cols serpentine path
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
            cell.setAttribute('aria-label', `ช่อง ${num}`);

            let iconHtml = '';
            if (num === 1) {
                cell.classList.add('cell-start');
                iconHtml = '🏕️';
            } else if (num === 80) {
                cell.classList.add('cell-finish');
                iconHtml = '🏰';
            } else if (config[num]) {
                const sp = config[num];
                cell.classList.add(`cell-${sp.type}`);
                if (sp.type === 'treasure') iconHtml = '📦';
                else if (sp.type === 'forward') iconHtml = '🚀';
                else if (sp.type === 'trap') iconHtml = '🕳️';
                else if (sp.type === 'water') iconHtml = '🌊';
                else if (sp.type === 'ghost') iconHtml = '👻';
                else if (sp.type === 'warp') iconHtml = '🌀';
                else if (sp.type === 'key') iconHtml = '🔑';
                else if (sp.type === 'door') iconHtml = '🚪';
            }

            cell.innerHTML = `<span class="cell-num">${num}</span> <span class="cell-icon">${iconHtml}</span> <div class="pawns-holder" id="pawns-holder-${num}"></div>`;
            grid.appendChild(cell);
        }
    }
}

// Update Gameplay UI & Render Pawns
function updateGameUI() {
    if (!gameState) return;

    const playersArr = Object.entries(gameState.players);
    const currentTurnKey = playersArr[gameState.turnIndex][0];
    const currentTurnPlayer = playersArr[gameState.turnIndex][1];

    // Status bar update
    document.getElementById('game-status-bar').textContent = `ถึงเทิร์นของ: ${currentTurnPlayer.animal.icon} ${currentTurnPlayer.name}`;

    // Update Player Cards
    const cardsContainer = document.getElementById('player-status-cards');
    cardsContainer.innerHTML = '';
    playersArr.forEach(([pId, p]) => {
        const card = document.createElement('div');
        card.className = `p-status-card ${pId === currentTurnKey ? 'active-turn' : ''}`;
        card.innerHTML = `
            <div><strong>${p.animal.icon} ${p.name}</strong></div>
            <div>ช่อง ${p.pos}/80</div>
            <div>🔑 ${p.keys} | 🛡️ ${p.armor}</div>
        `;
        cardsContainer.appendChild(card);
    });

    // Render Pawns on board
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

    // Roll button authority check
    const rollBtn = document.getElementById('btn-roll-dice');
    if (currentTurnKey === myPlayerId && !currentTurnPlayer.isBot) {
        rollBtn.disabled = false;
    } else {
        rollBtn.disabled = true;
    }

    // Host Bot Turn Handler
    if (isHost && currentTurnPlayer.isBot && gameState.status === 'playing') {
        setTimeout(() => {
            executeTurnRoll(currentTurnKey);
        }, 1800);
    }
}

// Handle Roll Dice Button Click
window.handleRollDice = function() {
    document.getElementById('btn-roll-dice').disabled = true;
    playSynthSound('roll');
    executeTurnRoll(myPlayerId);
};

// Execute Turn Roll & Movement Logic
function executeTurnRoll(pId) {
    const playersArr = Object.entries(gameState.players);
    const pData = gameState.players[pId];
    const diceRoll = Math.floor(Math.random() * 6) + 1;

    document.getElementById('dice-visual').textContent = diceRoll;
    announceSR(`${pData.name} ทอยลูกเต๋าได้ ${diceRoll}`);

    // Calculate Bounce Back Position
    let rawPos = pData.pos + diceRoll;
    let finalPos = rawPos;
    if (rawPos > 80) {
        finalPos = 80 - (rawPos - 80);
    }

    announceSR(`${pData.name} เดินจากช่อง ${pData.pos} ไปช่อง ${finalPos}`);
    playSynthSound('move');

    // Update position in Firebase
    const updates = {};
    updates[`games/Adventure80/rooms/${currentRoomId}/players/${pId}/pos`] = finalPos;

    // Check Landing Rules
    if (finalPos === 80) {
        // WINNER!
        playSynthSound('win');
        updates[`games/Adventure80/rooms/${currentRoomId}/status`] = 'ended';
        updates[`games/Adventure80/rooms/${currentRoomId}/winnerId`] = pId;
        update(ref(db), updates);
        return;
    }

    // Resolve Special Space (if applicable)
    const sp = gameState.boardConfig ? gameState.boardConfig[finalPos] : null;

    if (sp) {
        if (sp.type === 'treasure') {
            if (pData.armor < 3) {
                pData.armor += 1;
                updates[`games/Adventure80/rooms/${currentRoomId}/players/${pId}/armor`] = pData.armor;
                announceSR(`ช่อง ${finalPos} เป็นกล่องสมบัติ คุณได้รับเกราะศักดิ์สิทธิ์ 1 ชิ้น`);
                playSynthSound('treasure');
            }
        } else if (sp.type === 'key') {
            pData.keys += 1;
            updates[`games/Adventure80/rooms/${currentRoomId}/players/${pId}/keys`] = pData.keys;
            announceSR(`คุณพบกุญแจ ได้รับกุญแจ 1 ดอก ตอนนี้มีกุญแจ ${pData.keys} ดอก`);
            playSynthSound('key');
        } else if (sp.type === 'door') {
            if (pData.keys > 0) {
                if (!pData.isBot && pId === myPlayerId) {
                    showModal('คุณพบประตู!', 'คุณมีกุญแจ ต้องการใช้กุญแจเปิดประตูไปข้างหน้าหรือไม่?', 'ใช้กุญแจ', 'ไม่ใช้', () => {
                        pData.keys -= 1;
                        updates[`games/Adventure80/rooms/${currentRoomId}/players/${pId}/keys`] = pData.keys;
                        updates[`games/Adventure80/rooms/${currentRoomId}/players/${pId}/pos`] = sp.dest;
                        announceSR(`เปิดประตูสำเร็จ! วาร์ปไปยังช่อง ${sp.dest}`);
                        advanceTurn(updates);
                    }, () => {
                        advanceTurn(updates);
                    });
                    return;
                } else if (pData.isBot) {
                    pData.keys -= 1;
                    updates[`games/Adventure80/rooms/${currentRoomId}/players/${pId}/keys`] = pData.keys;
                    updates[`games/Adventure80/rooms/${currentRoomId}/players/${pId}/pos`] = sp.dest;
                }
            } else {
                announceSR(`คุณพบประตู แต่ไม่มีกุญแจ ประตูยังเปิดไม่ได้`);
            }
        } else if (['trap', 'water', 'ghost', 'warp'].includes(sp.type)) {
            const isBad = sp.val < 0 || sp.type === 'trap' || sp.type === 'water' || sp.type === 'ghost';
            if (isBad && pData.armor > 0) {
                if (!pData.isBot && pId === myPlayerId) {
                    showModal('พบเหตุการณ์อันตราย!', `ช่อง ${finalPos} มีเหตุการณ์ร้าย ใช้เกราะศักดิ์สิทธิ์ป้องกันหรือไม่?`, 'ใช้เกราะ', 'ไม่ใช้', () => {
                        pData.armor -= 1;
                        updates[`games/Adventure80/rooms/${currentRoomId}/players/${pId}/armor`] = pData.armor;
                        announceSR(`ใช้เกราะศักดิ์สิทธิ์ ป้องกันผลเสียสำเร็จ!`);
                        advanceTurn(updates);
                    }, () => {
                        applyBadEffect(pId, sp, updates);
                        advanceTurn(updates);
                    });
                    return;
                } else if (pData.isBot) {
                    pData.armor -= 1;
                    updates[`games/Adventure80/rooms/${currentRoomId}/players/${pId}/armor`] = pData.armor;
                    advanceTurn(updates);
                    return;
                }
            } else if (isBad) {
                applyBadEffect(pId, sp, updates);
            } else {
                // Positive Forward/Warp
                let nPos = Math.min(79, Math.max(1, finalPos + sp.val));
                updates[`games/Adventure80/rooms/${currentRoomId}/players/${pId}/pos`] = nPos;
                announceSR(`โชคดี! ได้เคลื่อนที่ไปยังช่อง ${nPos}`);
            }
        }
    }

    advanceTurn(updates);
}

// Apply Bad Effects helper
function applyBadEffect(pId, sp, updates) {
    const pData = gameState.players[pId];
    let nPos = Math.max(1, pData.pos + (sp.val || -3));
    updates[`games/Adventure80/rooms/${currentRoomId}/players/${pId}/pos`] = nPos;
    announceSR(`เกิดเหตุการณ์ร้าย ถอยหลังไปอยู่ช่อง ${nPos}`);
    playSynthSound('bad_event');
}

// Advance Turn Handler
function advanceTurn(updates) {
    const pKeys = Object.keys(gameState.players);
    const nextTurn = (gameState.turnIndex + 1) % pKeys.length;
    updates[`games/Adventure80/rooms/${currentRoomId}/turnIndex`] = nextTurn;
    update(ref(db), updates);
}

// Accessible Modal Dialog Component
function showModal(title, desc, confirmText, cancelText, onConfirm, onCancel) {
    const overlay = document.getElementById('modal-overlay');
    document.getElementById('modal-title').textContent = title;
    document.getElementById('modal-desc').textContent = desc;

    const btn1 = document.getElementById('btn-modal-action-1');
    const btn2 = document.getElementById('btn-modal-action-2');

    btn1.textContent = confirmText;
    btn2.textContent = cancelText;

    overlay.style.display = 'flex';
    document.getElementById('modal-title').focus();

    btn1.onclick = () => {
        overlay.style.display = 'none';
        if (onConfirm) onConfirm();
    };
    btn2.onclick = () => {
        overlay.style.display = 'none';
        if (onCancel) onCancel();
    };
}

// Show Result Screen
function showResultScreen() {
    switchScreen('screen-result', 'result-title');
    const winnerId = gameState.winnerId;
    const winner = gameState.players[winnerId];

    document.getElementById('result-winner-text').textContent = `🎉 ${winner.animal.icon} ${winner.name} เข้าสู่ช่อง 80 เป็นคนแรก!`;
    announceSR(`เกมจบแล้ว ${winner.name} เป็นผู้ชนะ!`);

    const rankingsBox = document.getElementById('result-rankings');
    rankingsBox.innerHTML = '<h3>อันดับการเดินทาง:</h3>';

    const sorted = Object.values(gameState.players).sort((a, b) => b.pos - a.pos);
    sorted.forEach((p, idx) => {
        rankingsBox.innerHTML += `<p>${idx + 1}. ${p.animal.icon} ${p.name} - ช่อง ${p.pos}/80</p>`;
    });
}

// Leave Room / Return to Main Menu
window.leaveRoom = function() {
    if (currentRoomId && myPlayerId) {
        remove(ref(db, `games/Adventure80/rooms/${currentRoomId}/players/${myPlayerId}`));
    }
    myPlayerId = null;
    currentRoomId = null;
    isHost = false;
    gameState = null;
    switchScreen('screen-welcome');
};

window.returnToLobbyOrMain = function() {
    window.leaveRoom();
};

// Initialize Application on DOM Ready
document.addEventListener('DOMContentLoaded', () => {
    initRoomListListener();
});