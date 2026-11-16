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

// Audio System
const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
const unlockAudio = () => { if (audioCtx.state === 'suspended') audioCtx.resume(); };
document.addEventListener('click', unlockAudio, { capture: true });
document.addEventListener('keydown', unlockAudio, { capture: true });

function playAudio(name) {
    try {
        fetch(`audio/${name}`)
            .then(res => res.arrayBuffer())
            .then(buf => audioCtx.decodeAudioData(buf))
            .then(decoded => {
                const source = audioCtx.createBufferSource();
                source.buffer = decoded;
                source.connect(audioCtx.destination);
                source.start(0);
            }).catch(() => {});
    } catch(e){}
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
    
    document.getElementById('btn-close-clue').onclick = () => document.getElementById('clue-modal').style.display = 'none';
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
    playAudio('select.mp3');
    runTransaction(ref(db, 'games/Escape/room_counter'), (cur) => (cur || 0) + 1).then((res) => {
        if (res.committed) {
            currentRoomId = 'Escape' + String(res.snapshot.val()).padStart(5, '0');
            myId = 'p1';
            isHost = true;
            const roomRef = ref(db, `games/Escape/rooms/${currentRoomId}`);
            onDisconnect(roomRef).remove(); // Host หลุด -> ลบห้อง
            set(roomRef, {
                status: 'waiting',
                players: { p1: createPlayerObj(myName, false) },
                botCount: 0, turnIndex: 0,
                log: { msg: '', ts: 0 }
            });
            enterRoomUI();
            announceSR(`สร้างห้องสำเร็จ รหัส ${currentRoomId}`);
        }
    });
}

function joinRoom(rId) {
    currentRoomId = rId;
    isHost = false;
    const roomRef = ref(db, `games/Escape/rooms/${currentRoomId}`);
    runTransaction(roomRef, (rData) => {
        if (!rData || rData.status !== 'waiting') return;
        const keys = Object.keys(rData.players || {});
        if (keys.length >= 4) return;
        let slot = 'p1';
        for(let i=1; i<=4; i++) { if(!keys.includes('p'+i)) { slot = 'p'+i; break; } }
        if (!rData.players) rData.players = {};
        rData.players[slot] = createPlayerObj(myName, false);
        myId = slot;
        return rData;
    }).then((res) => {
        if (res.committed && myId) {
            playAudio('select.mp3');
            onDisconnect(ref(db, `games/Escape/rooms/${currentRoomId}/players/${myId}`)).update({
                isBot: true, name: "บอท_" + myName
            });
            enterRoomUI();
            announceSR(`เข้าร่วมห้อง ${currentRoomId} สำเร็จ`);
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
    if (currentRoomId && myId) {
        if (isHost) remove(ref(db, `games/Escape/rooms/${currentRoomId}`));
        else update(ref(db, `games/Escape/rooms/${currentRoomId}/players/${myId}`), { isBot: true, name: "บอท_"+myName });
    }
    myId = null; currentRoomId = null; isHost = false; gameState = null;
    if(roomListener) roomListener(); // unsubscribe
    switchScreen('screen-welcome', 'welcome-title');
    announceSR("ออกจากห้องแล้ว");
};

function createPlayerObj(name, isBot) {
    return {
        name: name, isBot: isBot, floor: 6, hp: 200, status: 'normal',
        infection: 0, vaccine: 0, food: 0, armor: 0, gun: 0, ammo: 0,
        keys: 0, ghostKeys: 0, secretPassage: false,
        cluesFound: {},
        canUseDoor: false
    };
}

// ปรับปรุงการเพิ่ม-ลดบอท: เสถียร ไม่หลุดโฟกัสบน VoiceOver
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
        let slot = 'p1';
        for(let i=1; i<=4; i++) { if(!keys.includes('p'+i)) { slot = 'p'+i; break; } }
        update(ref(db, `games/Escape/rooms/${currentRoomId}`), {
            [`players/${slot}`]: createPlayerObj(`บอท${bCount+1}`, true),
            botCount: bCount + 1
        }).then(() => {
            playAudio('select.mp3');
            announceSR(`เพิ่มบอทสำเร็จ ขณะนี้มีบอท ${bCount + 1} ตัว รวม ${total + 1} คน`, 'polite');
        });
    } else if(delta < 0) {
        if (bCount <= 0) {
            announceSR("ไม่มีบอทให้ลดแล้ว", 'assertive');
            return;
        }
        const bots = keys.filter(k=>gameState.players[k].isBot);
        if(bots.length > 0) {
            const updates = { botCount: bCount - 1 };
            updates[`players/${bots[bots.length-1]}`] = null;
            update(ref(db, `games/Escape/rooms/${currentRoomId}`), updates).then(() => {
                playAudio('select.mp3');
                announceSR(`ลดบอทสำเร็จ ขณะนี้มีบอท ${bCount - 1} ตัว รวม ${total - 1} คน`, 'polite');
            });
        }
    }
}

// Game Logic
function generateRoomsForFloor(floorNum) {
    let arr = [
        {type:'key'}, {type:'ghostKey'}, {type:'secret'}, {type:'clue', cid:0}, {type:'clue', cid:1}, {type:'clue', cid:2},
        {type:'zombie'}, {type:'zombie'}, {type:'zombie'}, {type:'zombie'},
        {type:'vaccine'}, {type:'vaccine'}, {type:'vaccine'},
        {type:'food'}, {type:'food'}, {type:'food'},
        {type:'armor'}, {type:'puzzle'}, {type:'gun'}, {type:'battery'}
    ];
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
    
    for(let i=pKeys.length-1; i>0; i--) { const j = Math.floor(Math.random()*(i+1)); [pKeys[i], pKeys[j]] = [pKeys[j], pKeys[i]]; }
    
    const floorsData = {};
    for(let f=6; f>=1; f--) floorsData[f] = generateRoomsForFloor(f);

    update(ref(db, `games/Escape/rooms/${currentRoomId}`), {
        status: 'playing', turnIndex: 0, order: pKeys,
        floors: floorsData, log: { msg: 'เกมเริ่มขึ้นแล้ว หาทางลงไปยังชั้น 1 และหนีออกจากตึก!', ts: Date.now() }
    });
}

let localTs = 0;
function setupRoomListener() {
    roomListener = onValue(ref(db, `games/Escape/rooms/${currentRoomId}`), (snap) => {
        gameState = snap.val();
        if(!gameState) { announceSR("ห้องถูกปิด", "assertive"); setTimeout(()=>window.location.href='../index.html', 2000); return; }
        
        if(gameState.status === 'waiting') {
            document.getElementById('lobby-room-title').textContent = `ห้อง: ${currentRoomId}`;
            const pList = Object.values(gameState.players||{});
            const totalPlayers = pList.length;
            document.getElementById('participant-count').textContent = `${totalPlayers}/4`;
            const listEl = document.getElementById('participant-list');
            listEl.innerHTML = '';
            pList.forEach(p => {
                listEl.innerHTML += `<div class="participant-card">${p.name} ${p.isBot?'(บอท)':''}</div>`;
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
                
                // Dimming อย่างสมบูรณ์ทั้ง attribute และ VoiceOver Accessibility Label
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
            if(!document.getElementById('screen-game').classList.contains('active-screen')) {
                switchScreen('screen-game', 'game-floor-title');
            }
            
            // ลำดับการประกาศ: ประกาศผลลัพธ์เหตุการณ์ที่เกิดขึ้นก่อน
            if(gameState.log && gameState.log.ts > localTs) {
                localTs = gameState.log.ts;
                if(gameState.log.msg) {
                    announceSR(gameState.log.msg, 'polite');
                }
                document.getElementById('latest-event-display').textContent = gameState.log.msg;
            }
            
            // อัปเดตหน้าจอและการประกาศเทิร์นถัดไป
            renderGameUI();
        } else if (gameState.status === 'ended') {
            if(!document.getElementById('screen-result').classList.contains('active-screen')) {
                switchScreen('screen-result', 'result-title');
                renderResult();
            }
        }
    });
}

function renderGameUI() {
    if(!gameState) return;
    const myP = gameState.players[myId];
    if(!myP) return;
    
    document.getElementById('game-floor-title').textContent = `คุณอยู่ชั้น ${myP.floor}`;
    
    const curPId = gameState.order[gameState.turnIndex];
    const curP = gameState.players[curPId];

    // หากตัวละครอยู่ในสถานะพักฟื้นครบ 1 เทิร์นแล้ว เมื่อถึงเทิร์นของตนเองจะได้รับการฟื้นฟู HP เป็น 100 และกลับสู่สถานะปกติเพื่อเล่นต่อได้
    if(curP.status === 'resting') {
        if(isHost) {
            curP.hp = 100;
            curP.status = 'normal';
            emitLog(`${curP.name} พักฟื้นเสร็จสิ้น ฟื้น HP เป็น 100 และพร้อมลุยต่อ`);
            commitPlayer(curPId, curP);
        }
        return;
    }
    
    const isMyTurn = (curPId === myId && !myP.isBot && myP.status === 'normal');
    
    // ประกาศเทิร์นและหน่วงเวลาการคิดของบอทให้สมจริง (2.2 วินาที)
    if(lastAnnouncedTurn !== `${gameState.turnIndex}_${curPId}`) {
        lastAnnouncedTurn = `${gameState.turnIndex}_${curPId}`;
        announceSR(`ถึงเทิร์นของ ${curP.name}`);
        playAudio('turn.mp3');
        if(isHost && curP.isBot && curP.status === 'normal') {
            setTimeout(() => executeBotTurn(curPId), 2200);
        } else if(curP.status !== 'normal' && isHost) {
            setTimeout(() => endTurn(curPId), 1200);
        }
    }
    
    document.getElementById('game-status-bar').textContent = `เทิร์นของ: ${curP.name}`;
    
    // แถบสถานะส่วนตัว (แสดงอาหารในคลังด้วย)
    const statusText = `สถานะของคุณ: ${myP.name} ชั้น ${myP.floor} HP ${myP.hp} อาหารในคลัง ${myP.food} วัคซีน ${myP.vaccine} กุญแจ ${myP.keys} ปืน ${myP.gun} กระสุน ${myP.ammo}`;
    const statusContainer = document.getElementById('my-player-status');
    statusContainer.setAttribute('aria-label', statusText);
    statusContainer.innerHTML = `<span aria-hidden="true">${myP.name} ชั้น ${myP.floor} | HP: ${myP.hp} | อาหาร: ${myP.food} | วัคซีน: ${myP.vaccine} | กุญแจ: ${myP.keys} | ปืน: ${myP.gun}(${myP.ammo})</span>`;
    
    // Status Cards (สำหรับสายตา)
    const cardsEl = document.getElementById('player-status-cards');
    cardsEl.innerHTML = '';
    gameState.order.forEach(pid => {
        const p = gameState.players[pid];
        let cClass = 'p-status-card';
        if(pid === curPId) cClass += ' active-turn';
        if(p.status === 'dead') cClass += ' dead';
        if(p.status === 'escaped') cClass += ' escaped';
        cardsEl.innerHTML += `<div class="${cClass}" aria-hidden="true">
            <strong>${p.name} (ชั้น ${p.floor})</strong><br>
            HP: ${p.hp} | ติดเชื้อ: ${p.infection}<br>
            กุญแจ: ${p.keys} | อาหาร: ${p.food} | เกราะ: ${p.armor} | ปืน: ${p.gun}(${p.ammo})
        </div>`;
    });
    
    // Board (Grid 20 ห้อง)
    const boardEl = document.getElementById('escape-board');
    boardEl.innerHTML = '';
    const floorData = gameState.floors[myP.floor];
    if(floorData) {
        for(let i=0; i<20; i++) {
            const roomData = floorData[i];
            const btn = document.createElement('button');
            btn.className = 'board-cell';
            const num = i + 1;
            btn.textContent = num.toString();
            
            if(roomData.explored) {
                btn.classList.add('explored');
                if (roomData.type === 'clue') {
                    btn.setAttribute('aria-label', `ห้องคำใบ้ ${num} สำรวจแล้ว สามารถกดอ่านคำใบ้ซ้ำได้`);
                    if (isMyTurn) {
                        btn.onclick = () => exploreRoom(i, roomData);
                    } else {
                        btn.disabled = true;
                        btn.setAttribute('aria-disabled', 'true');
                    }
                } else {
                    btn.disabled = true;
                    btn.setAttribute('aria-disabled', 'true');
                    btn.setAttribute('aria-label', `ห้อง ${num} สำรวจแล้ว`);
                }
            } else {
                btn.setAttribute('aria-label', `ห้อง ${num} กดเพื่อสำรวจ`);
                if(isMyTurn) {
                    btn.onclick = () => exploreRoom(i, roomData);
                } else {
                    btn.disabled = true;
                    btn.setAttribute('aria-disabled', 'true');
                }
            }
            boardEl.appendChild(btn);
        }
    }
    
    // บันได หรือ ประตูหนีออกจากตึก
    const btnStairs = document.getElementById('btn-use-stairs');
    const isFloor1 = (myP.floor === 1);
    
    if (isFloor1) {
        btnStairs.textContent = "🚪 เปิดประตูหนีออกจากตึก";
    } else {
        btnStairs.textContent = "🚪 ลงบันไดไปชั้นล่าง";
    }
    
    const canUseStairs = isMyTurn && (myP.keys > 0 || myP.ghostKeys > 0 || myP.secretPassage || myP.canUseDoor);
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
    gameState.order.forEach(pid => {
        const p = gameState.players[pid];
        let txt = `${p.name} - ชั้น ${p.floor} | HP: ${p.hp}`;
        if(p.status !== 'normal') txt += ` (${p.status})`;
        txt += ` | ติดเชื้อ: ${p.infection} | อาหาร: ${p.food} | วัคซีน: ${p.vaccine} | กุญแจ: ${p.keys} | เกราะ: ${p.armor} | ปืน: ${p.gun} | กระสุน: ${p.ammo}`;
        list.innerHTML += `<div class="participant-card">${txt}</div>`;
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

// Turn Processing
async function processTurnStartEffects(pId) {
    let p = gameState.players[pId];
    if(p.status === 'dead' || p.status === 'escaped') return p;
    
    if(p.status === 'resting') {
        p.hp = 100;
        p.status = 'normal';
        await emitLog(`${p.name} พักฟื้นเสร็จสิ้น ฟื้น HP เป็น 100 และพร้อมลุยต่อ`);
        return p;
    }
    
    // 1. จัดการการติดเชื้อ & วัคซีน
    if(p.infection > 0) {
        if(p.vaccine > 0) {
            p.vaccine--; p.infection = 0;
            await emitLog(`${p.name} ใช้วัคซีนรักษาการติดเชื้ออัตโนมัติ`);
        } else {
            p.hp -= p.infection;
            await emitLog(`${p.name} เสีย ${p.infection} HP จากการติดเชื้อ (เหลือ ${p.hp} HP)`);
        }
    }
    
    // 2. จัดการอาหารในคลังอัตโนมัติ: หาก HP ต่ำกว่า 200 และมีอาหาร จะกินรักษาอัตโนมัติ 1 ชิ้น (+30 HP)
    if(p.hp > 0 && p.hp < 200 && p.food > 0) {
        const heal = Math.min(30, 200 - p.hp);
        p.food--;
        p.hp += heal;
        await emitLog(`${p.name} กินอาหารจากคลังอัตโนมัติ ฟื้นฟู HP ${heal} หน่วย (HP ปัจจุบัน ${p.hp}, เหลืออาหาร ${p.food} ชิ้น)`);
    }
    
    // 3. ตรวจสอบการหมดสติ/เสียชีวิต
    if(p.hp <= 0) {
        p.status = 'dead';
        await emitLog(`${p.name} เสียชีวิตแล้ว...`);
    } else if (p.hp <= 60 && p.status === 'normal') {
        p.status = 'resting';
        await emitLog(`${p.name} HP ต่ำกว่า 60 ต้องเข้าสู่สถานะพักฟื้นในเทิร์นนี้`);
    }
    
    return p;
}

async function exploreRoom(roomIdx, roomData) {
    if(!gameState) return;
    const pId = myId;
    let p = await processTurnStartEffects(pId);
    if(p.status !== 'normal') {
        await commitPlayer(pId, p);
        setTimeout(() => endTurn(pId), 1500);
        return;
    }

    if(roomData.type === 'clue') {
        openClueModal(roomData.clueId);
        return;
    }
    
    await update(ref(db), { [`games/Escape/rooms/${currentRoomId}/floors/${p.floor}/${roomIdx}/explored`]: true });
    
    let msg = `${p.name} สำรวจห้อง ${roomIdx+1}: `;
    
    switch(roomData.type) {
        case 'key': p.keys++; msg += "พบกุญแจปกติ 1 ดอก"; break;
        case 'ghostKey': p.ghostKeys++; msg += "พบกุญแจผี 1 ดอก"; break;
        case 'secret': p.secretPassage = true; msg += "พบทางลงลับ! สามารถลงข้ามชั้นได้"; break;
        case 'vaccine': p.vaccine++; msg += "พบวัคซีน 1 ขวด เก็บเข้าคลัง"; break;
        case 'food': 
            if(p.hp < 200) {
                const heal = Math.min(30, 200 - p.hp);
                p.hp += heal;
                msg += `พบอาหาร กินทันทีฟื้นฟู HP ${heal} หน่วย (HP ปัจจุบัน ${p.hp})`;
            } else {
                p.food = (p.food || 0) + 1;
                msg += `พบอาหาร แต่ HP เต็ม 200 แล้ว จึงเก็บอาหารเข้าคลัง (มีอาหารสะสม ${p.food} ชิ้น)`;
            }
            break;
        case 'armor': p.armor++; msg += "พบเกราะสวรรค์ 1 ชิ้น"; break;
        case 'battery': 
            if(p.ammo < 2) { p.ammo++; msg += "พบแบตเตอรี่ ได้กระสุน 1 นัด"; }
            else { msg += "พบแบตเตอรี่ แต่กระสุนเต็มแล้ว"; }
            break;
        case 'gun':
            if(p.gun === 0) { p.gun = 1; if(p.ammo < 2) p.ammo++; msg += "พบปืนเลเซอร์และกระสุน 1 นัด!"; }
            else if(p.ammo < 2) { p.ammo++; msg += "พบปืนเลเซอร์ ซ้ำ แปลงเป็นกระสุน 1 นัด"; }
            else { msg += "พบปืนเลเซอร์ ซ้ำ แต่กระสุนเต็มแล้ว"; }
            break;
        case 'puzzle': msg += "พบเหตุการณ์พิเศษ: ไม่มีอะไรเกิดขึ้น (ปลอดภัย)"; break;
        case 'zombie':
            msg += "พบซอมบี้! ";
            if(p.gun > 0 && p.ammo > 0) {
                p.ammo--; msg += "ใช้ปืนยิงป้องกันสำเร็จ เสียกระสุน 1 นัด";
            } else if (p.armor > 0) {
                p.armor--; msg += "เกราะสวรรค์ป้องกันการโจมตีสำเร็จ เกราะพัง";
            } else {
                if(p.infection === 0) p.infection = 20;
                else p.infection = 40;
                msg += `ถูกโจมตี ติดเชื้อระดับ ${p.infection} HP/เทิร์น`;
            }
            break;
    }
    
    await emitLog(msg);
    await commitPlayer(pId, p);
    // หน่วงเวลา 1.5 วินาทีเพื่อให้ VoiceOver อ่านเหตุการณ์ครบก่อนเปลี่ยนเทิร์น
    setTimeout(() => endTurn(pId), 1500);
}

function openClueModal(cid) {
    const clue = window.EscapeTowerClueBank ? window.EscapeTowerClueBank.getClueById(cid) : null;
    if(!clue) {
        emitLog(`${gameState.players[myId].name} สำรวจห้องคำใบ้ แต่ระบบคำใบ้ไม่พร้อมใช้งาน รับสิทธิ์ผ่านประตูไปเลย`);
        gameState.players[myId].canUseDoor = true;
        commitPlayer(myId, gameState.players[myId]).then(() => setTimeout(()=>endTurn(myId), 1500));
        return;
    }
    
    const m = document.getElementById('clue-modal');
    document.getElementById('clue-question').textContent = clue.question;
    
    const inputEl = document.getElementById('clue-answer-input');
    inputEl.value = '';
    inputEl.dataset.cid = cid;
    
    document.getElementById('clue-feedback').textContent = '';
    m.style.display = 'flex';
    document.getElementById('clue-title').focus();
    announceSR(clue.question);
}

function submitClue() {
    const inputEl = document.getElementById('clue-answer-input');
    const cid = inputEl.dataset.cid;
    const ans = inputEl.value.trim();
    if(!ans) return;
    
    const isCorrect = window.EscapeTowerClueBank.checkAnswer(cid, ans);
    if(isCorrect) {
        document.getElementById('clue-modal').style.display = 'none';
        const p = gameState.players[myId];
        p.canUseDoor = true;
        const targetDesc = (p.floor === 1) ? "เปิดประตูทางออกตึก" : `เปิดประตูชั้น ${p.floor}`;
        emitLog(`${p.name} ตอบคำใบ้ถูกต้อง ได้รับสิทธิ์${targetDesc}`);
        commitPlayer(myId, p).then(() => setTimeout(()=>endTurn(myId), 1500));
    } else {
        document.getElementById('clue-feedback').textContent = "คำตอบผิด ลองใหม่ในเทิร์นหน้า";
        announceSR("คำตอบผิด ลองใหม่ในเทิร์นหน้า");
        setTimeout(() => {
            document.getElementById('clue-modal').style.display = 'none';
            emitLog(`${gameState.players[myId].name} ตอบคำใบ้ผิด ประตูยังไม่เปิด`);
            setTimeout(() => endTurn(myId), 1200);
        }, 1500);
    }
}

// ลงบันได หรือ เปิดประตูหนีออกจากตึก
async function useStairs() {
    let p = await processTurnStartEffects(myId);
    if(p.status !== 'normal') { 
        await commitPlayer(myId, p); 
        setTimeout(() => endTurn(myId), 1500); 
        return; 
    }

    let used = "";
    if(p.secretPassage) { p.secretPassage = false; used = "ทางลับ"; }
    else if(p.canUseDoor) { p.canUseDoor = false; used = "รหัสคำใบ้"; }
    else if(p.ghostKeys > 0) { p.ghostKeys--; used = "กุญแจผี"; }
    else if(p.keys > 0) { p.keys--; used = "กุญแจปกติ"; }

    // กรณีอยู่ชั้น 1: การใช้กุญแจ/รหัส จะเป็นการเปิดประตูหนีออกจากตึกทันที (ชนะเกม!)
    if(p.floor === 1) {
        p.status = 'escaped';
        p.canUseDoor = false;
        let msg = `${p.name} ใช้${used} ปลดล็อกประตูทางออกชั้น 1 และหนีรอดออกจากตึกได้สำเร็จ เป็นผู้ชนะ!!`;
        await emitLog(msg);
        await commitPlayer(myId, p);
        await triggerEndGame(myId);
        return;
    }

    // กรณีอยู่ชั้น 2-6: ลงบันไดไปชั้นล่าง
    let down = (used === "ทางลับ") ? 2 : 1;
    p.floor = Math.max(1, p.floor - down);
    p.canUseDoor = false;
    
    let msg = `${p.name} ใช้${used} ลงมายังชั้น ${p.floor}`;
    if(p.floor === 1) {
        msg += " ถึงชั้น 1 แล้ว! จงหากุญแจหรือรหัสเพื่อปลดล็อกประตูหนีออกจากตึก!";
    }
    
    await emitLog(msg);
    await commitPlayer(myId, p);
    
    setTimeout(() => endTurn(myId), 1500);
}

// ระบบบอทเล่น: เพิ่มการคิด + ระบบเก็บอาหาร + กติกาชั้น 1
async function executeBotTurn(botId) {
    let p = gameState.players[botId];
    if(!p || p.status !== 'normal') { 
        setTimeout(() => endTurn(botId), 1200); 
        return; 
    }
    
    p = await processTurnStartEffects(botId);
    if(p.status !== 'normal') {
        await commitPlayer(botId, p);
        setTimeout(() => endTurn(botId), 1500);
        return;
    }
    
    const canUnlock = (p.secretPassage || p.ghostKeys > 0 || p.keys > 0 || p.canUseDoor);
    
    // บอทอยู่ชั้น 1 แล้วมีไอเท็มเปิดประตูหนี
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
    
    // บอทอยู่ชั้น 2-6 แล้วมีไอเท็มลงชั้นล่าง
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
        await emitLog(msg);
        await commitPlayer(botId, p);
        setTimeout(() => endTurn(botId), 1500);
        return;
    }
    
    // สำรวจห้องบนชั้นปัจจุบัน
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
        switch(roomData.type) {
            case 'key': p.keys++; msg += "พบกุญแจปกติ"; break;
            case 'ghostKey': p.ghostKeys++; msg += "พบกุญแจผี"; break;
            case 'secret': p.secretPassage = true; msg += "พบทางลับ"; break;
            case 'vaccine': p.vaccine++; msg += "พบวัคซีน"; break;
            case 'food': 
                if(p.hp < 200) {
                    p.hp = Math.min(200, p.hp + 30);
                    msg += "กินอาหารฟื้น HP";
                } else {
                    p.food = (p.food || 0) + 1;
                    msg += "เก็บอาหารเข้าคลังสำรอง";
                }
                break;
            case 'armor': p.armor++; msg += "ได้เกราะ"; break;
            case 'battery': if(p.ammo<2) p.ammo++; msg += "ได้กระสุน"; break;
            case 'gun': if(p.gun===0) p.gun=1; if(p.ammo<2) p.ammo++; msg += "ได้ปืนและกระสุน"; break;
            case 'puzzle': msg += "ไม่เกิดอะไรขึ้น"; break;
            case 'zombie':
                msg += "เจอซอมบี้! ";
                if(p.gun>0 && p.ammo>0) { p.ammo--; msg+="ยิงป้องกันสำเร็จ"; }
                else if(p.armor>0) { p.armor--; msg+="เกราะกันสำเร็จ"; }
                else { p.infection = p.infection===0?20:40; msg+="ติดเชื้อ!"; }
                break;
        }
        await emitLog(msg);
        await commitPlayer(botId, p);
    } else {
        await emitLog(`${p.name} ไม่เหลือห้องให้สำรวจ ข้ามเทิร์น`);
    }
    setTimeout(() => endTurn(botId), 1500);
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

async function emitLog(msg) {
    await update(ref(db), { [`games/Escape/rooms/${currentRoomId}/log`]: { msg, ts: Date.now() } });
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

    const order = gameState.order;
    let idx = gameState.turnIndex;
    idx = (idx + 1) % order.length;
    
    update(ref(db), { [`games/Escape/rooms/${currentRoomId}/turnIndex`]: idx });
}

function renderResult() {
    const winnerName = gameState.winner ? gameState.players[gameState.winner].name : "ไม่มีผู้รอดชีวิต";
    announceSR(`เกมจบลงแล้ว ผู้ชนะคือ ${winnerName}`, "assertive");
    
    document.getElementById('result-title').textContent = `ผู้ชนะ: ${winnerName}`;
    
    const resEl = document.getElementById('result-list');
    resEl.innerHTML = '';
    const pArr = Object.values(gameState.players);
    pArr.forEach(p => {
        let txt = `${p.name} - ชั้น ${p.floor} `;
        if(p.status === 'escaped') txt += "(หนีรอดสำเร็จ! 🏆)";
        else if(p.status === 'dead') txt += "(เสียชีวิต 💀)";
        else txt += `(ติดอยู่ในตึก HP: ${p.hp})`;
        resEl.innerHTML += `<div class="participant-card">${txt}</div>`;
    });
}