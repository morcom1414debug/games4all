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

// Name and Navigation
window.confirmNameAndEnterLobby = function() {
    const nameInput = document.getElementById('player-name-input');
    myPlayerName = nameInput.value.trim() || 'ผู้เล่นใหม่';
    switchScreen('screen-lobby', 'lobby-heading');
    announceSR(`ยินดีต้อนรับ ${myPlayerName} เข้าสู่ล็อบบี้การผจญภัย`);
};

window.leaveRoomAndGoHome = function() {
    if (currentRoomId) window.leaveRoom();
    switchScreen('screen-welcome', 'welcome-title');
};

window.toggleManual = function() {
    const m = document.getElementById('manual-section');
    if (m.style.display === 'none') {
        m.style.display = 'block';
        m.focus();
        announceSR("เปิดคู่มือการเล่นแล้ว");
    } else {
        m.style.display = 'none';
        announceSR("ปิดคู่มือการเล่นแล้ว");
    }
};

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
    if (!myPlayerName) myPlayerName = 'ผู้เล่น 1';

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
            announceSR(`สร้างห้องสำเร็จ รหัสห้องคือ ${currentRoomId}`);
        }
    });
};

// Join Room Handler
window.joinAdventureRoom = function(rId) {
    if (!myPlayerName) myPlayerName = 'ผู้เล่นเข้าใหม่';

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
            // Client disconnecting turns them into a bot
            onDisconnect(ref(db, `games/Adventure80/rooms/${currentRoomId}/players/${myPlayerId}`)).update({
                isBot: true,
                name: "บอท" + myPlayerName
            });
            
            document.getElementById('lobby-menu-section').style.display = 'none';
            document.getElementById('lobby-room-section').style.display = 'block';
            
            setupRoomListener();
            announceSR(`เข้าร่วมห้อง ${currentRoomId} เรียบร้อยแล้ว`);
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
            announceSR("เจ้าของห้องออกจากเกม ห้องถูกปิด", 'assertive');
            setTimeout(() => {
                window.location.href = '../index.html';
            }, 3000);
            return;
        }

        // Check for disconnected players converting to bots
        if (gameState.players) {
            for (const pId in gameState.players) {
                if (previousPlayersState[pId] && !previousPlayersState[pId].isBot && gameState.players[pId].isBot) {
                    announceSR(`${previousPlayersState[pId].name} หลุดจากห้อง แปลงเป็นบอท`, 'polite');
                }
            }
            previousPlayersState = JSON.parse(JSON.stringify(gameState.players));
        }

        updateLobbyUI();

        // Synced Actions Announcements Listener
        if (gameState.lastAction && gameState.lastAction.ts > localLastActionTs) {
            localLastActionTs = gameState.lastAction.ts;
            announceSR(gameState.lastAction.msg, 'polite');
        }

        if (gameState.status === 'playing') {
            if (document.getElementById('screen-game').classList.contains('active-screen') === false) {
                showStartGameAnimation();
            } else {
                updateGameUI();
            }
        } else if (gameState.status === 'ended') {
            showResultScreen();
        }
    });
}

function showStartGameAnimation() {
    const overlay = document.getElementById('anim-start-overlay');
    overlay.style.display = 'flex';
    announceSR('การผจญภัยเริ่มต้นขึ้นแล้ว เตรียมตัวให้พร้อม!', 'assertive');
    
    setTimeout(() => {
        overlay.style.display = 'none';
        switchScreen('screen-game', 'game-status-bar');
        renderBoardGrid();
        updateGameUI();
    }, 2500);
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
    if (!isHost) return;
    lastAnnouncedTurnKey = null;
    const boardConfig = generateBoardConfiguration();
    const roomRef = ref(db, `games/Adventure80/rooms/${currentRoomId}`);

    const pKeys = Object.keys(gameState.players);
    for (let i = pKeys.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [pKeys[i], pKeys[j]] = [pKeys[j], pKeys[i]];
    }

    update(roomRef, {
        status: 'playing',
        boardConfig,
        turnIndex: 0,
        playerOrder: pKeys
    });
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
            let accessibleName = `ช่อง ${num}`;

            if (num === 1) {
                cell.classList.add('cell-start');
                iconHtml = '🏕️';
                accessibleName += ' จุดเริ่มต้น';
            } else if (num === 80) {
                cell.classList.add('cell-finish');
                iconHtml = '🏰';
                accessibleName += ' เส้นชัยปราสาท';
            } else if (config[num]) {
                const sp = config[num];
                cell.classList.add(`cell-${sp.type}`);
                
                if (sp.type === 'rest') { iconHtml = '⛺'; accessibleName += ' จุดพักผ่อน ไม่มีเหตุการณ์'; }
                else if (sp.type === 'treasure') { iconHtml = '📦'; accessibleName += ' หีบสมบัติ'; }
                else if (sp.type === 'forward') { iconHtml = '🚀'; accessibleName += ' วาร์ป'; }
                else if (sp.type === 'trap') { iconHtml = '🕳️'; accessibleName += ' หลุมพราง'; }
                else if (sp.type === 'water') { iconHtml = '🌊'; accessibleName += ' น้ำเชี่ยว'; }
                else if (sp.type === 'ghost') { iconHtml = '👻'; accessibleName += ' ผีหลอก'; }
                else if (sp.type === 'warp') { iconHtml = '🌀'; accessibleName += ' ไซโคลน'; }
                else if (sp.type === 'key') { iconHtml = '🔑'; accessibleName += ' กล่องลึกลับ'; }
                else if (sp.type === 'door') { iconHtml = '🚪'; accessibleName += ` ประตูทางลัด ไปช่อง ${sp.dest}`; }
                else if (sp.type === 'bonus') { iconHtml = '🎁'; accessibleName += ' ลาภลอย'; }
                else if (sp.type === 'secret') { iconHtml = '🔮'; accessibleName += ' เหตุการณ์ลับ'; }
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

    document.getElementById('game-status-bar').textContent = `ถึงเทิร์นของ: ${currentTurnPlayer.animal.icon} ${currentTurnPlayer.name}`;

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
            const turnMsg = `ถึงเทิร์นของ ${currentTurnPlayer.name}`;
            announceSR(turnMsg);
            
            // Focus the roll button when it is the local player's turn, after announcement
            if (currentTurnKey === myPlayerId && !currentTurnPlayer.isBot && !rollBtn.disabled) {
                const duration = Math.max(1500, turnMsg.length * 50) + 200;
                setTimeout(() => {
                    const btn = document.getElementById('btn-roll-dice');
                    if (btn && !btn.disabled) {
                        btn.focus();
                    }
                }, duration);
            }
        }
    }

    const cardsContainer = document.getElementById('player-status-cards');
    cardsContainer.innerHTML = '';
    
    // Group all statuses into a single readable object for Screen Readers
    let allStatusText = "สถานะผู้เล่นทั้งหมด: ";
    playersArr.forEach(([pId, p], index) => {
        allStatusText += `${p.name} ช่อง ${p.pos} กุญแจ ${p.keys} เกราะ ${p.armor}`;
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
            <div>ช่อง ${p.pos} กุญแจ ${p.keys} เกราะ ${p.armor}</div>
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

    // Host Bot Turn Engine
    if (isHost && currentTurnPlayer.isBot && gameState.status === 'playing' && !gameState.turnExecuting) {
        // Lock executing flag
        update(ref(db), { [`games/Adventure80/rooms/${currentRoomId}/turnExecuting`]: true });
        setTimeout(() => { executeTurnAsync(currentTurnKey, true); }, 1500);
    }
}

// Emits an action and pauses briefly so all clients process audio sequentially
async function syncActionEmit(msg) {
    const ts = Date.now() + Math.random();
    await update(ref(db), { [`games/Adventure80/rooms/${currentRoomId}/lastAction`]: { msg, ts } });
    await delayAsync(2200); // Wait long enough for Screen Readers to speak out the sequence
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
    document.getElementById('btn-roll-dice').disabled = true;
    if (myTurnTimer) {
        clearTimeout(myTurnTimer);
        myTurnTimer = null;
    }
    update(ref(db), { [`games/Adventure80/rooms/${currentRoomId}/turnExecuting`]: true });
    playSynthSound('roll');
    executeTurnAsync(myPlayerId, isAuto);
};

// Event Chain Engine (Core Game Logic Refactored)
async function executeTurnAsync(pId, isAuto = false) {
    const pData = gameState.players[pId];
    const autoMode = pData.isBot || isAuto;
    const diceRoll = Math.floor(Math.random() * 6) + 1;
    document.getElementById('dice-visual').textContent = diceRoll;
    
    await syncActionEmit(`${pData.name} ทอยลูกเต๋าได้ ${diceRoll}`);
    
    let currentPos = pData.pos;
    let targetPos = currentPos + diceRoll;
    if (targetPos > 80) targetPos = 80 - (targetPos - 80); // Bounce back rule

    playSynthSound('move');
    await syncActionEmit(`${pData.name} เดินจากช่อง ${currentPos} ไปยังช่อง ${targetPos}`);
    
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
        if (sp.type === 'rest') typeNameTH = 'จุดพักผ่อน';
        else if (sp.type === 'treasure') typeNameTH = 'หีบสมบัติ';
        else if (sp.type === 'forward') typeNameTH = 'วาร์ป';
        else if (sp.type === 'trap') typeNameTH = 'หลุมพราง';
        else if (sp.type === 'water') typeNameTH = 'น้ำเชี่ยว';
        else if (sp.type === 'ghost') typeNameTH = 'ผีหลอก';
        else if (sp.type === 'warp') typeNameTH = 'ไซโคลน';
        else if (sp.type === 'key') typeNameTH = 'กล่องลึกลับ';
        else if (sp.type === 'door') typeNameTH = 'ประตูทางลัด';
        else if (sp.type === 'bonus') typeNameTH = 'ลาภลอย';
        else if (sp.type === 'secret') typeNameTH = 'เหตุการณ์ลับ';

        await syncActionEmit(`${pData.name}ตกช่อง ${currentPos} เป็นช่อง${typeNameTH}`);

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
                playSynthSound('treasure');
                await syncActionEmit(`พบหีบสมบัติ เปิดหีบ พบเกราะศักดิ์สิทธิ์ ได้รับเกราะศักดิ์สิทธิ์ 1 ชิ้น`);
            } else {
                await syncActionEmit(`พบหีบสมบัติ แต่หีบถูกเปิดไปแล้ว เหลือเพียงเศษฝุ่น ไม่ได้อะไรเลย`);
            }
            break; // Event chain ends as player did not change location
        } else if (sp.type === 'key') {
            if (sp.charges > 0) {
                sp.charges--;
                await update(ref(db), { [`games/Adventure80/rooms/${currentRoomId}/boardConfig/${currentPos}/charges`]: sp.charges });
                pData.keys++;
                await syncStateDB(pId, { keys: pData.keys });
                playSynthSound('key');
                await syncActionEmit(`พบกล่องลึกลับ เปิดกล่อง พบกุญแจโบราณ ได้รับกุญแจโบราณ 1 ดอก`);
            } else {
                await syncActionEmit(`พบกล่องลึกลับ แต่กล่องถูกเปิดไปแล้ว เหลือเพียงเศษฝุ่น ไม่ได้อะไรเลย`);
            }
            break; // Does not change location
        } else if (sp.type === 'door') {
            const doorNames = ["ประตูมิติ", "ประตูกาลเวลา", "ประตูศิลาโบราณ", "ประตูเวทมนตร์", "ประตูลับ"];
            const dName = doorNames[Math.floor(Math.random() * doorNames.length)];
            if (pData.keys > 0) {
                let useKey = true;
                if (!autoMode && pId === myPlayerId) {
                    useKey = await showModalAsync(`พบ${dName}!`, 'คุณมีกุญแจ ต้องการใช้กุญแจเปิดประตูไปข้างหน้าหรือไม่?', 'ใช้กุญแจ', 'ไม่ใช้');
                }
                if (useKey) {
                    pData.keys--;
                    await syncStateDB(pId, { keys: pData.keys });
                    await syncActionEmit(`ใช้กุญแจเปิด${dName}สำเร็จ วาร์ปไปยังช่อง ${sp.dest}`);
                    currentPos = sp.dest;
                    movedToNewSpace = true;
                } else {
                    await syncActionEmit(`เลือกที่จะไม่ใช้กุญแจ ประตูยังคงปิดอยู่`);
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
                    useArmor = await showModalAsync('พบอันตราย!', `ช่อง ${currentPos} มีเหตุการณ์ร้าย ใช้เกราะศักดิ์สิทธิ์ป้องกันหรือไม่?`, 'ใช้เกราะป้องกัน', 'ไม่ใช้');
                }
                if (useArmor) {
                    pData.armor--;
                    await syncStateDB(pId, { armor: pData.armor });
                    await syncActionEmit(`ใช้เกราะศักดิ์สิทธิ์ป้องกันผลเสียสำเร็จ ปลอดภัยแล้ว!`);
                    break;
                } else {
                    playSynthSound('bad_event');
                    await syncActionEmit(`รับผลร้าย ถอยหลัง ${Math.abs(randVal)} ช่อง`);
                    currentPos = Math.max(1, currentPos + randVal);
                    movedToNewSpace = true;
                }
            } else if (isBad) {
                playSynthSound('bad_event');
                await syncActionEmit(`เกิดเหตุการณ์ร้าย ถอยหลัง ${Math.abs(randVal)} ช่อง`);
                currentPos = Math.max(1, currentPos + randVal);
                movedToNewSpace = true;
            } else {
                await syncActionEmit(`โชคดี เดินหน้าเพิ่ม ${randVal} ช่อง`);
                let targetPos = currentPos + randVal;
                if (targetPos > 80) targetPos = 80 - (targetPos - 80);
                currentPos = targetPos;
                movedToNewSpace = true;
            }
        } else if (sp.type === 'forward') {
            let randVal = Math.floor(Math.random() * 10) + 3;
            await syncActionEmit(`พลังพิเศษ เดินหน้า ${randVal} ช่อง`);
            let targetPos = currentPos + randVal;
            if (targetPos > 80) targetPos = 80 - (targetPos - 80);
            currentPos = targetPos;
            movedToNewSpace = true;
        } else if (sp.type === 'bonus') {
            const randEffect = Math.floor(Math.random() * 3);
            if (randEffect === 0) {
                pData.keys++;
                await syncStateDB(pId, { keys: pData.keys });
                playSynthSound('key');
                await syncActionEmit(`ลาภลอย! ได้รับกุญแจ 1 ชิ้น`);
                break;
            } else if (randEffect === 1) {
                pData.armor++;
                await syncStateDB(pId, { armor: pData.armor });
                playSynthSound('treasure');
                await syncActionEmit(`ลาภลอย! ได้รับเกราะศักดิ์สิทธิ์ 1 ชิ้น`);
                break;
            } else {
                let randVal = Math.floor(Math.random() * 10) + 3;
                let actionDesc = `เดินหน้า ${randVal} ช่อง`;
                await syncActionEmit(`ลาภลอย! เกิดการวาร์ป ${actionDesc}`);
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
            await syncActionEmit(`พบเหตุการณ์ลับ: ${sName} ${actionDesc}`);
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
    playSynthSound('win');
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

    document.getElementById('result-winner-text').textContent = `🎉 ${winner.animal.icon} ${winner.name} เข้าสู่ช่อง 80 เป็นคนแรก!`;
    announceSR(`การผจญภัยสิ้นสุดลง ผู้ชนะคือ ${winner.name} !`, 'assertive');

    const rankingsBox = document.getElementById('result-rankings');
    rankingsBox.innerHTML = '<h3>อันดับการเดินทาง:</h3>';

    const sorted = Object.values(gameState.players).sort((a, b) => b.pos - a.pos);
    sorted.forEach((p) => {
        rankingsBox.innerHTML += `<p>${p.name} ช่อง ${p.pos}</p>`;
    });
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
    myPlayerId = null;
    currentRoomId = null;
    isHost = false;
    gameState = null;
    lastAnnouncedTurnKey = null;
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

// Initialize Application on DOM Ready
document.addEventListener('DOMContentLoaded', () => {
    initRoomListListener();

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

        nameInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                if (nameInput.value.trim().length > 0) {
                    window.confirmNameAndEnterLobby();
                }
            }
        });
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
                let allStatusText = "สถานะผู้เล่นทั้งหมด: ";
                const pOrder = gameState.playerOrder || Object.keys(gameState.players);
                pOrder.forEach((pId, index) => {
                    const p = gameState.players[pId];
                    allStatusText += `${p.name} ช่อง ${p.pos} กุญแจ ${p.keys} เกราะ ${p.armor}`;
                    if (index < pOrder.length - 1) allStatusText += ", ";
                });
                announceSR(allStatusText, 'polite');
            }
        }
    });
});