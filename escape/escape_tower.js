import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
import { getDatabase, ref, set, onValue, update, runTransaction, remove, onDisconnect } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-database.js";

// Exact Firebase Configuration สำหรับ Escape Tower
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

// Animal Avatars Pool
const ANIMAL_POOL = [
    { icon: '🦊', name: 'จิ้งจอก' }, { icon: '🐼', name: 'แพนด้า' },
    { icon: '🐰', name: 'กระต่าย' }, { icon: '🐨', name: 'โคอาลา' },
    { icon: '🐯', name: 'เสือ' }, { icon: '🦁', name: 'สิงโต' }
];

// Global Game States
let myPlayerName = '';
let myPlayerId = null;
let currentRoomId = null;
let isHost = false;
let gameState = null;
let roomListener = null;
let localLastActionTs = 0;
let speechQueue = [];
let isSpeaking = false;
let lastAnnouncedTurnKey = null;
let previousPlayersState = {};
let hintsData = null; // เก็บข้อมูลคำใบ้

// Audio System States
let bgmSource = null;
let audioBuffers = {};
let isShowingWinnerScene = false;

// Audio Context Web API (Web Audio API)
const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
const unlockAudio = () => {
    if (audioCtx.state === 'suspended') audioCtx.resume();
};
document.addEventListener('click', unlockAudio, { capture: true });
document.addEventListener('touchstart', unlockAudio, { capture: true });
document.addEventListener('keydown', unlockAudio, { capture: true });

async function loadAudio(name) {
    if (audioBuffers[name]) return audioBuffers[name];
    try {
        const response = await fetch(`audio/${name}`);
        const arrayBuffer = await response.arrayBuffer();
        const audioBuffer = await audioCtx.decodeAudioData(arrayBuffer);
        audioBuffers[name] = audioBuffer;
        return audioBuffer;
    } catch (e) {
        console.warn('Audio load error or missing file:', name);
        return null;
    }
}

function playAudio(name, loop = false) {
    return new Promise(async (resolve) => {
        if (audioCtx.state === 'suspended') await audioCtx.resume();
        const buffer = await loadAudio(name);
        if (!buffer) {
            resolve();
            return;
        }
        const source = audioCtx.createBufferSource();
        source.buffer = buffer;
        source.connect(audioCtx.destination);
        source.loop = loop;
        source.onended = () => resolve(source);
        source.start(0);

        if (loop && name === 'bgm.mp3') {
            if (bgmSource) { try { bgmSource.stop(); } catch (e) { } }
            bgmSource = source;
        }
    });
}

function stopBGM() {
    if (bgmSource) {
        try { bgmSource.stop(); } catch (e) { }
        bgmSource = null;
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
                const cleanupFocus = () => {
                    h.removeAttribute('tabindex');
                    h.removeEventListener('blur', cleanupFocus);
                };
                h.addEventListener('blur', cleanupFocus);
            }
        }
    }
}

// Fetch Hints
function fetchHints() {
    if (typeof window.EscapeTowerClueBank !== 'undefined') {
        hintsData = window.EscapeTowerClueBank;
    } else {
        hintsData = null;
    }
}

// Board Random Generator สำหรับ Escape Tower
function generateTowerFloors() {
    const floors = {};
    for (let f = 1; f <= 6; f++) {
        const roomTypes = [
            'key_normal', 'key_skeleton', 'secret_stair',
            'hint_1', 'hint_2', 'hint_3',
            'zombie', 'zombie', 'zombie', 'zombie',
            'vaccine', 'vaccine', 'vaccine',
            'food', 'food', 'food',
            'armor', 'puzzle', 'laser_gun', 'battery'
        ];
        // Shuffle
        for (let i = roomTypes.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [roomTypes[i], roomTypes[j]] = [roomTypes[j], roomTypes[i]];
        }
        
        // สุ่มคำใบ้ประจำชั้น 3 ข้อจากคลังคำใบ้ต้นฉบับ
        const selectedClues = (typeof window.EscapeTowerClueBank !== 'undefined') 
            ? window.EscapeTowerClueBank.selectCluesForGame(f) 
            : [];

        let hintCount = 0;
        const rooms = {};

        for(let i=0; i<20; i++) {
            const type = roomTypes[i];
            let clueData = null;

            // ถ้าสุ่มได้ห้องคำใบ้ ให้แนบข้อมูลคำใบ้ประจำห้องไปด้วย
            if (type.startsWith('hint_')) {
                clueData = selectedClues[hintCount] || null;
                hintCount++;
            }

            rooms[i] = { 
                type: type, 
                exploredBy: [],
                clue: clueData // บันทึกคำใบ้ลงในห้อง
            };
        }
        floors[f] = { rooms: rooms, doorsOpened: [] }; // doorsOpened เก็บ player ids ที่เปิดประตูลงชั้นแล้ว
    }
    return floors;
}

window.createEscapeRoom = async function () {
    const nameInput = document.getElementById('player-name-input');
    if(nameInput) myPlayerName = nameInput.value.trim() || 'ผู้เล่น 1';
    else myPlayerName = 'ผู้เล่น 1';

    playAudio('select.mp3');
    await fetchHints(); // โหลดคำใบ้ก่อนสร้างห้อง

    const counterRef = ref(db, 'games/Escape/room_counter');
    runTransaction(counterRef, (cur) => (cur || 0) + 1).then((res) => {
        if (res.committed) {
            currentRoomId = 'Escape' + String(res.snapshot.val()).padStart(5, '0');
            myPlayerId = 'p1';
            isHost = true;

            const roomRef = ref(db, `games/Escape/rooms/${currentRoomId}`);
            onDisconnect(roomRef).remove();

            const initialPlayers = {
                p1: { 
                    name: myPlayerName, animal: ANIMAL_POOL[0], isBot: false, 
                    floor: 6, hp: 200, infectionRate: 0, 
                    inventory: { food: [], vaccines: 0, keys: 0, skeletonKeys: 0, armor: 0, ammo: 0, hasGun: false },
                    status: 'active'
                }
            };

            set(roomRef, {
                status: 'waiting',
                players: initialPlayers,
                botCount: 0,
                turnIndex: 0,
                towerConfig: null,
                lastAction: { msg: '', ts: 0 }
            });

            switchScreen('screen-lobby', 'lobby-heading');
            setupRoomListener();
            announceSR(`สร้างห้องสำเร็จ รหัสห้องคือ ${currentRoomId}`);
        }
    });
};

window.joinEscapeRoom = async function (rId) {
    const nameInput = document.getElementById('player-name-input');
    if(nameInput) myPlayerName = nameInput.value.trim() || 'ผู้เข้าใหม่';
    else myPlayerName = 'ผู้เข้าใหม่';

    currentRoomId = (rId || document.getElementById('room-id-input').value).trim();
    if(!currentRoomId) return;

    isHost = false;
    await fetchHints();

    const roomRef = ref(db, `games/Escape/rooms/${currentRoomId}`);
    runTransaction(roomRef, (roomData) => {
        if (!roomData || roomData.status !== 'waiting') return;
        const pKeys = Object.keys(roomData.players || {});
        if (pKeys.length >= 4) return; // รวมบอทไม่เกิน 4

        let nextSlot = 'p1';
        for (let i = 1; i <= 4; i++) {
            if (!pKeys.includes('p' + i)) {
                nextSlot = 'p' + i;
                break;
            }
        }

        const usedAnimals = Object.values(roomData.players).map(p => p.animal.icon);
        const availAnimals = ANIMAL_POOL.filter(a => !usedAnimals.includes(a.icon));

        if (!roomData.players) roomData.players = {};
        roomData.players[nextSlot] = {
            name: myPlayerName, animal: availAnimals[0] || ANIMAL_POOL[0], isBot: false, 
            floor: 6, hp: 200, infectionRate: 0, 
            inventory: { food: [], vaccines: 0, keys: 0, skeletonKeys: 0, armor: 0, ammo: 0, hasGun: false },
            status: 'active'
        };

        myPlayerId = nextSlot;
        return roomData;
    }).then((res) => {
        if (res.committed) {
            playAudio('select.mp3');
            onDisconnect(ref(db, `games/Escape/rooms/${currentRoomId}/players/${myPlayerId}`)).update({
                isBot: true, name: "บอท" + myPlayerName
            });
            switchScreen('screen-lobby', 'lobby-heading');
            setupRoomListener();
            announceSR(`เข้าร่วมห้อง ${currentRoomId} เรียบร้อยแล้ว`);
        } else {
            announceSR("ไม่สามารถเข้าห้องได้ ห้องอาจเต็มหรือเริ่มเกมไปแล้ว", 'assertive');
        }
    });
};

window.adjustBot = function(delta) {
    if (!isHost || !gameState) return;
    const curBots = gameState.botCount || 0;
    const humanCount = Object.values(gameState.players).filter(p => !p.isBot).length;
    const newBots = curBots + delta;

    if (newBots < 0 || (humanCount + newBots) > 4) return;

    const roomRef = ref(db, `games/Escape/rooms/${currentRoomId}`);
    const updatedPlayers = { ...gameState.players };

    if (delta > 0) {
        const botSlot = 'p' + (humanCount + newBots);
        const botNames = ['บอทผู้รอด 1', 'บอทผู้รอด 2', 'บอทผู้รอด 3'];
        const usedAnimals = Object.values(updatedPlayers).map(p => p.animal.icon);
        const availAnimals = ANIMAL_POOL.filter(a => !usedAnimals.includes(a.icon));

        updatedPlayers[botSlot] = {
            name: botNames[newBots - 1] || `บอท ${newBots}`,
            animal: availAnimals[0] || ANIMAL_POOL[1],
            isBot: true, floor: 6, hp: 200, infectionRate: 0,
            inventory: { food: [], vaccines: 0, keys: 0, skeletonKeys: 0, armor: 0, ammo: 0, hasGun: false },
            status: 'active'
        };
    } else {
        const botSlots = Object.keys(updatedPlayers).filter(k => updatedPlayers[k].isBot);
        if (botSlots.length > 0) {
            delete updatedPlayers[botSlots[botSlots.length - 1]];
        }
    }

    update(roomRef, { botCount: newBots, players: updatedPlayers });
};

function setupRoomListener() {
    const roomRef = ref(db, `games/Escape/rooms/${currentRoomId}`);
    roomListener = onValue(roomRef, (snapshot) => {
        gameState = snapshot.val();
        if (!gameState) {
            announceSR("ห้องถูกปิดแล้ว กลับสู่หน้าแรก", 'assertive');
            setTimeout(() => { location.reload(); }, 2000);
            return;
        }

        updateLobbyUI();

        if (gameState.lastAction && gameState.lastAction.ts > localLastActionTs) {
            localLastActionTs = gameState.lastAction.ts;
            announceSR(gameState.lastAction.msg, 'polite');
            if (gameState.lastAction.audio) playAudio(gameState.lastAction.audio);
        }

        if (gameState.status === 'playing') {
            if (!document.getElementById('screen-game').classList.contains('active-screen')) {
                switchScreen('screen-game', 'game-title');
                playAudio('bgm.mp3', true);
            }
            updateGameUI();
            checkAndRunBotTurn();
        }
    });
}

function updateLobbyUI() {
    if(!gameState) return;
    const pList = Object.values(gameState.players || {});
    const container = document.getElementById('participant-list');
    if(container) {
        container.innerHTML = '';
        pList.forEach(p => {
            const div = document.createElement('div');
            div.className = 'room-list li'; // จำลอง class ตาม css
            div.textContent = `${p.animal.icon} ${p.name} ${p.isBot ? '(บอท)' : ''}`;
            container.appendChild(div);
        });
    }

    const startBtn = document.getElementById('btn-start-game');
    if(startBtn) {
        if(isHost) {
            startBtn.style.display = 'block';
            startBtn.disabled = (gameState.status !== 'waiting');
        } else {
            startBtn.style.display = 'none';
        }
    }
}

window.startEscapeGame = function() {
    if (!isHost || !gameState || gameState.status !== 'waiting') return;
    const towerConfig = generateTowerFloors();
    const pKeys = Object.keys(gameState.players || {});
    
    for (let i = pKeys.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [pKeys[i], pKeys[j]] = [pKeys[j], pKeys[i]];
    }

    update(ref(db, `games/Escape/rooms/${currentRoomId}`), {
        status: 'playing',
        towerConfig,
        turnIndex: 0,
        playerOrder: pKeys
    });
};

function checkAndRunBotTurn() {
    if (!isHost || !gameState || gameState.status !== 'playing' || gameState.turnExecuting) return;
    const pOrder = gameState.playerOrder;
    const currentTurnKey = pOrder[gameState.turnIndex];
    const currentTurnPlayer = gameState.players[currentTurnKey];
    
    if (currentTurnPlayer && currentTurnPlayer.isBot && currentTurnPlayer.status === 'active') {
        update(ref(db), { [`games/Escape/rooms/${currentRoomId}/turnExecuting`]: true });
        setTimeout(() => { executeBotTurn(currentTurnKey); }, 2000);
    }
}

function updateGameUI() {
    if (!gameState) return;
    const me = gameState.players[myPlayerId];
    if(!me) return;

    // อัปเดต UI ส่วนตัวและทุกคน
    const floorIndicator = document.getElementById('floor-description');
    if(floorIndicator) floorIndicator.textContent = `คุณอยู่ชั้น ${me.floor}`;

    const statusBar = document.getElementById('latest-event');
    if(statusBar) {
        statusBar.textContent = `HP: ${me.hp} | ติดเชื้อ: ${me.infectionRate > 0 ? 'ใช่(-'+me.infectionRate+'/เทิร์น)' : 'ไม่'} | กุญแจ: ${me.inventory.keys} | กระสุน: ${me.inventory.ammo} / ปืน: ${me.inventory.hasGun ? 'มี':'ไม่มี'}`;
    }

    renderRoomGrid(me.floor);
    
    const pOrder = gameState.playerOrder || [];
    const currentTurnKey = pOrder[gameState.turnIndex];
    const currentTurnPlayer = gameState.players[currentTurnKey];

    if (currentTurnPlayer) {
        const turnUniqueId = `${currentRoomId}_${gameState.turnIndex}_${currentTurnKey}`;
        if (lastAnnouncedTurnKey !== turnUniqueId) {
            lastAnnouncedTurnKey = turnUniqueId;
            announceSR(`ถึงเทิร์นของ ${currentTurnPlayer.name} (อยู่ชั้น ${currentTurnPlayer.floor})`);
            playAudio('turn.mp3');
            
            // ประมวลผลเทิร์นตัวเองก่อนอนุญาตให้กด
            if (currentTurnKey === myPlayerId && !currentTurnPlayer.isBot) {
                processMyTurnStart(currentTurnPlayer);
            }
        }
    }
}

function renderRoomGrid(floorNum) {
    const grid = document.getElementById('room-grid');
    if(!grid || !gameState.towerConfig) return;
    grid.innerHTML = '';
    
    const currentTurnKey = gameState.playerOrder[gameState.turnIndex];
    const isMyTurn = (currentTurnKey === myPlayerId && !gameState.turnExecuting);
    const floorData = gameState.towerConfig[floorNum];

    for(let i=0; i<20; i++) {
        const room = floorData.rooms[i];
        const btn = document.createElement('button');
        btn.className = 'room-button';
        
        let isLocked = room.exploredBy.includes('all');
        if (room.type.startsWith('hint_')) {
            isLocked = false; // ห้องคำใบ้เข้าได้เรื่อยๆ
        }

        btn.textContent = isLocked ? 'สำรวจแล้ว' : `ห้อง ${i+1}`;
        btn.disabled = isLocked || !isMyTurn;
        btn.onclick = () => window.exploreRoom(i);
        
        // ถ้าเป็นห้องคำใบ้ที่เคยเปิดโดยตัวเอง ให้บอกว่าเป็นคำใบ้อะไร
        if (room.exploredBy.includes(myPlayerId) && room.type.startsWith('hint_')) {
            btn.textContent += ' (อ่านคำใบ้)';
        }
        
        grid.appendChild(btn);
    }
}

async function emitAction(msg, audio = null) {
    const ts = Date.now() + Math.random();
    await update(ref(db), { [`games/Escape/rooms/${currentRoomId}/lastAction`]: { msg, ts, audio } });
}

async function processMyTurnStart(pData) {
    update(ref(db), { [`games/Escape/rooms/${currentRoomId}/turnExecuting`]: true });
    
    let hp = pData.hp;
    let infRate = pData.infectionRate;
    let status = pData.status;

    if (hp <= 0) {
        await emitAction(`${pData.name} เสียชีวิตแล้ว ข้ามเทิร์น`);
        status = 'dead';
        await endTurnFlow({ status });
        return;
    }

    if (infRate > 0) {
        hp -= infRate;
        await emitAction(`${pData.name} เสีย HP ${infRate} จากการติดเชื้อ (เหลือ HP ${hp})`, 'hurt.mp3');
    }

    // ฟื้นฟูอัตโนมัติ (ข้ามเทิร์น 1 ครั้ง)
    if (hp > 0 && hp <= 60 && status !== 'recovering') {
        await emitAction(`${pData.name} HP เหลือน้อย เข้าสู่สถานะพักฟื้น ข้าม 1 เทิร์น`);
        await endTurnFlow({ hp, status: 'recovering' });
        return;
    } else if (status === 'recovering') {
        hp = 100;
        await emitAction(`${pData.name} พักฟื้นเสร็จสิ้น ฟื้นฟู HP เป็น 100 พร้อมสู้ต่อ!`, 'heal.mp3');
        status = 'active'; // ฟื้นแล้วให้เล่นต่อในรอบหน้า (หรือรอบนี้เลย? ตามกติกาคือข้าม 1 เทิร์น แล้วฟื้นเป็น 100 ดังนั้นรอบถัดไปจะ active)
        // ตามกติกา: ข้ามเทิร์นของตัวเอง 1 ครั้งแล้วฟื้น
        await endTurnFlow({ hp, status });
        return;
    }

    // Auto use vaccine/food if needed logic should go here or when attacked.
    // For now, allow player to select room. Unlock turn Execution.
    if(hp <= 0) {
        await emitAction(`${pData.name} ทนพิษบาดแผลไม่ไหว เสียชีวิตแล้ว`);
        await endTurnFlow({ hp: 0, status: 'dead' });
        return;
    }

    await update(ref(db, `games/Escape/rooms/${currentRoomId}`), {
        [`players/${myPlayerId}/hp`]: hp,
        [`players/${myPlayerId}/status`]: status,
        'turnExecuting': false
    });
}

window.exploreRoom = async function(roomIndex) {
    if (gameState.turnExecuting) return;
    update(ref(db), { [`games/Escape/rooms/${currentRoomId}/turnExecuting`]: true });

    const me = gameState.players[myPlayerId];
    const floor = me.floor;
    const room = gameState.towerConfig[floor].rooms[roomIndex];
    let roomType = room.type;

    await emitAction(`${me.name} เลือกสำรวจ ห้องที่ ${roomIndex+1}`);

    let lockRoom = true;
    let updates = { ...me.inventory };
    let hpUpdate = me.hp;
    let infUpdate = me.infectionRate;

    switch(roomType) {
        case 'key_normal':
            updates.keys++;
            await emitAction(`พบกุญแจปกติ! เก็บไว้ไขประตูลงชั้น`, 'item.mp3');
            break;
        case 'key_skeleton':
            updates.skeletonKeys++;
            await emitAction(`พบกุญแจผี! ใช้ลงชั้นได้ทันที`, 'item.mp3');
            break;
        case 'secret_stair':
            await emitAction(`พบทางลงลับ! (สามารถใช้ข้ามชั้นได้ที่ประตู)`, 'door.mp3');
            // logic is handled when clicking go down stairs
            break;
        case 'hint_1':
        case 'hint_2':
        case 'hint_3':
            lockRoom = false; // ห้องคำใบ้กลับมาอ่านซ้ำได้ ไม่ล็อก
            if (room.clue) {
                await emitAction(`พบคำใบ้ที่ ${roomType.replace('hint_', '')}: "${room.clue.question}"`);
                
                // อัปเดตข้อความบนหน้าจอ puzzle-content เฉพาะเครื่องผู้เล่นตนเอง
                const puzzleBox = document.getElementById('puzzle-content');
                if (puzzleBox) {
                    puzzleBox.innerHTML = `
                        <p><strong>คำใบ้ประจำห้อง (ชั้น ${floor}):</strong></p>
                        <p>${room.clue.question}</p>
                    `;
                }
            } else {
                await emitAction(`พบห้องคำใบ้ แต่ไม่พบข้อมูลคำใบ้ในระบบ`);
            }
            break;
        case 'zombie':
            if (updates.hasGun && updates.ammo > 0) {
                updates.ammo--;
                await emitAction(`พบซอมบี้! ปืนเลเซอร์ยิงอัตโนมัติ กำจัดซอมบี้ปลอดภัย (เหลือกระสุน ${updates.ammo})`, 'laser.mp3');
            } else if (updates.armor > 0) {
                updates.armor--;
                await emitAction(`พบซอมบี้! โดนโจมตีแต่เกราะสวรรค์ป้องกันไว้ได้ เกราะพังลง`, 'shield.mp3');
            } else {
                if (infUpdate === 0) infUpdate = 20;
                else if (infUpdate === 20) infUpdate = 40;
                await emitAction(`ถูกซอมบี้โจมตี! ติดเชื้อระดับเสีย HP ${infUpdate}/เทิร์น`, 'zombie.mp3');
                
                // Auto use vaccine
                if (updates.vaccines > 0) {
                    updates.vaccines--;
                    infUpdate = 0;
                    await emitAction(`ใช้วัคซีนสำรองอัตโนมัติ รักษาการติดเชื้อหายเป็นปกติ`, 'heal.mp3');
                }
            }
            break;
        case 'vaccine':
            if (infUpdate > 0) {
                infUpdate = 0;
                await emitAction(`พบวัคซีน! นำมารักษาการติดเชื้อทันที`, 'heal.mp3');
            } else {
                updates.vaccines++;
                await emitAction(`พบวัคซีน! เก็บสำรองไว้ในคลัง`, 'item.mp3');
            }
            break;
        case 'food':
            if (hpUpdate < 200) {
                hpUpdate = Math.min(200, hpUpdate + 30); // ฟื้นกลางๆ 30
                await emitAction(`พบอาหาร กินทันที ฟื้นฟู HP เป็น ${hpUpdate}`, 'heal.mp3');
            } else if (updates.food.length < 3) {
                updates.food.push(30);
                await emitAction(`พบอาหาร แต่ HP เต็มแล้ว จึงเก็บเข้ากระเป๋า (สูงสุด 3 ชิ้น)`, 'item.mp3');
            } else {
                await emitAction(`พบอาหาร แต่กระเป๋าเต็มแล้ว จึงทิ้งไว้ที่เดิม`);
            }
            break;
        case 'armor':
            updates.armor++;
            await emitAction(`พบเกราะสวรรค์! ป้องกันการถูกโจมตีได้ 1 ครั้ง`, 'item.mp3');
            break;
        case 'puzzle':
            await emitAction(`พบห้องปริศนา! (ระบบกิจกรรมพิเศษ ไม่ส่งผลขัดแย้งกติกาหลัก)`);
            break;
        case 'laser_gun':
            updates.hasGun = true;
            if (updates.ammo < 2) updates.ammo++;
            await emitAction(`พบปืนเลเซอร์! ได้รับปืนพร้อมกระสุน (กระสุนรวม: ${updates.ammo}/2)`, 'item.mp3');
            break;
        case 'battery':
            if (updates.ammo < 2) updates.ammo++;
            await emitAction(`พบแบตเตอรี่กระสุนเลเซอร์ (กระสุนรวม: ${updates.ammo}/2)`, 'item.mp3');
            break;
    }

    const nextRoomState = room.exploredBy;
    if (lockRoom) nextRoomState.push('all');
    else if (!nextRoomState.includes(myPlayerId)) nextRoomState.push(myPlayerId);

    const dbUpdates = {};
    dbUpdates[`games/Escape/rooms/${currentRoomId}/towerConfig/${floor}/rooms/${roomIndex}/exploredBy`] = nextRoomState;
    
    await endTurnFlow({
        inventory: updates,
        hp: hpUpdate,
        infectionRate: infUpdate
    }, dbUpdates);
};

window.useStairs = async function() {
    if (gameState.turnExecuting) return;
    const me = gameState.players[myPlayerId];
    if(gameState.playerOrder[gameState.turnIndex] !== myPlayerId) return;
    
    // Check keys
    if (me.inventory.keys > 0 || me.inventory.skeletonKeys > 0) {
        update(ref(db), { [`games/Escape/rooms/${currentRoomId}/turnExecuting`]: true });
        
        let updates = { ...me.inventory };
        if (updates.skeletonKeys > 0) updates.skeletonKeys--;
        else updates.keys--;

        let nextFloor = me.floor - 1;
        await emitAction(`${me.name} ใช้กุญแจเปิดประตูลงชั้นสำเร็จ! ลงไปสู่ชั้น ${nextFloor}`, 'door.mp3');
        
        if (nextFloor === 0) {
            await emitAction(`🎉 ${me.name} หนีออกจากตึกสำเร็จ เป็นผู้ชนะ!`, 'win.mp3');
            await endTurnFlow({ floor: 0, status: 'escaped' }, { [`games/Escape/rooms/${currentRoomId}/status`]: 'ended' });
            return;
        }

        await endTurnFlow({ floor: nextFloor, inventory: updates }, {});
    } else {
        announceSR("คุณไม่มีกุญแจลงชั้น ต้องไขปริศนาหรือหาทางลับแทน", 'assertive');
    }
};

async function endTurnFlow(playerUpdates, extraDbUpdates = {}) {
    const pKeys = gameState.playerOrder;
    let nextTurn = (gameState.turnIndex + 1) % pKeys.length;
    
    const updates = {
        [`games/Escape/rooms/${currentRoomId}/turnIndex`]: nextTurn,
        [`games/Escape/rooms/${currentRoomId}/turnExecuting`]: false,
        ...extraDbUpdates
    };

    for(const [k, v] of Object.entries(playerUpdates)) {
        updates[`games/Escape/rooms/${currentRoomId}/players/${myPlayerId}/${k}`] = v;
    }

    await update(ref(db), updates);
}

async function executeBotTurn(botId) {
    const bot = gameState.players[botId];
    if (bot.status !== 'active') {
        await emitAction(`${bot.name} ไม่สามารถเล่นได้ในเทิร์นนี้`);
        await advanceBotTurn(botId, { status: bot.status });
        return;
    }

    // สำรวจห้องแบบสุ่มที่ยังไม่ถูกสำรวจ
    const floor = bot.floor;
    const floorRooms = gameState.towerConfig[floor].rooms;
    let unvisited = [];
    for(let i=0; i<20; i++) {
        if (!floorRooms[i].exploredBy.includes('all')) unvisited.push(i);
    }

    if (unvisited.length > 0) {
        const pick = unvisited[Math.floor(Math.random() * unvisited.length)];
        await emitAction(`${bot.name} เลือกสำรวจห้อง ${pick + 1}`);
        // Bot logic simulation simplified: Just lock room to keep flow moving.
        const dbUpdates = {};
        dbUpdates[`games/Escape/rooms/${currentRoomId}/towerConfig/${floor}/rooms/${pick}/exploredBy`] = ['all'];
        await advanceBotTurn(botId, {}, dbUpdates);
    } else {
        await emitAction(`${bot.name} ไม่มีห้องให้สำรวจแล้ว ข้ามเทิร์น`);
        await advanceBotTurn(botId, {});
    }
}

async function advanceBotTurn(botId, playerUpdates, extraDbUpdates = {}) {
    const pKeys = gameState.playerOrder;
    let nextTurn = (gameState.turnIndex + 1) % pKeys.length;
    const updates = {
        [`games/Escape/rooms/${currentRoomId}/turnIndex`]: nextTurn,
        [`games/Escape/rooms/${currentRoomId}/turnExecuting`]: false,
        ...extraDbUpdates
    };
    for(const [k, v] of Object.entries(playerUpdates)) {
        updates[`games/Escape/rooms/${currentRoomId}/players/${botId}/${k}`] = v;
    }
    await update(ref(db), updates);
}