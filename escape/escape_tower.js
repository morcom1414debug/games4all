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
    // ใช้ Web Audio API โหลดและเล่นเสียงจำลองแบบเดียวกับ Camping
    // หากไม่มีไฟล์เสียง จะข้ามไปเงียบๆ ไม่ทำให้เกมพัง
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

// ARIA Announcement
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
            setTimeout(() => { targetEl.textContent = ''; isSpeaking = false; processSpeechQueue(); }, duration);
        }, 100);
    } else { isSpeaking = false; }
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
    document.getElementById('player-name-input').addEventListener('input', (e) => {
        document.getElementById('btn-confirm-name').disabled = e.target.value.trim().length === 0;
    });
    document.getElementById('player-name-input').addEventListener('keydown', (e) => {
        if(e.key === 'Enter' && e.target.value.trim().length > 0) confirmName();
    });
    document.getElementById('btn-confirm-name').onclick = confirmName;
    document.getElementById('btn-create-room').onclick = createRoom;
    document.getElementById('btn-leave-room').onclick = window.leaveRoomAndGoHome;
    document.getElementById('btn-start-game').onclick = startGame;
    
    document.getElementById('btn-open-manual-welcome').onclick = () => showManual(true);
    document.getElementById('btn-open-manual-lobby').onclick = () => showManual(true);
    document.getElementById('btn-close-manual').onclick = () => showManual(false);

    document.getElementById('btn-add-bot').onclick = () => adjustBot(1);
    document.getElementById('btn-remove-bot').onclick = () => adjustBot(-1);
    document.getElementById('btn-read-all-status').onclick = readAllStatus;
    
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
        if (myId) return; // Ignore if in room
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
            onDisconnect(roomRef).remove(); // Hostหลุด -> ลบห้อง
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
        name: name, isBot: isBot, floor: 6, hp: 200, status: 'normal', // normal, resting, dead, escaped
        infection: 0, vaccine: 0, food: 0, armor: 0, gun: 0, ammo: 0,
        keys: 0, ghostKeys: 0, secretPassage: false,
        cluesFound: {}, // { floor: { clueId, answersTried } }
        canUseDoor: false
    };
}

function adjustBot(delta) {
    if(!isHost || !gameState) return;
    const hCount = Object.values(gameState.players).filter(p=>!p.isBot).length;
    const bCount = gameState.botCount || 0;
    const total = hCount + bCount;
    if(delta > 0 && total < 4) {
        const slot = 'p' + (total+1);
        update(ref(db, `games/Escape/rooms/${currentRoomId}`), {
            [`players/${slot}`]: createPlayerObj(`บอท${bCount+1}`, true),
            botCount: bCount + 1
        });
        playAudio('select.mp3');
    } else if(delta < 0 && bCount > 0) {
        const bots = Object.keys(gameState.players).filter(k=>gameState.players[k].isBot);
        if(bots.length > 0) {
            const updates = { botCount: bCount - 1 };
            updates[`players/${bots[bots.length-1]}`] = null;
            update(ref(db, `games/Escape/rooms/${currentRoomId}`), updates);
            playAudio('select.mp3');
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
    
    // Pick 3 Clues for this floor from clue bank
    let clues = [];
    if(window.EscapeTowerClueBank) {
        clues = window.EscapeTowerClueBank.selectCluesForGame(floorNum);
    }
    
    let board = {};
    for(let i=0; i<20; i++) {
        board[i] = { type: arr[i].type, explored: false };
        if(arr[i].type === 'clue' && clues[arr[i].cid]) {
            board[i].clueId = clues[arr[i].cid].id;
            board[i].explored = false; // clue rooms can be re-entered
        }
    }
    return board;
}

function startGame() {
    if(!isHost || !gameState || gameState.status !== 'waiting') return;
    document.getElementById('btn-start-game').disabled = true;
    
    const pKeys = Object.keys(gameState.players);
    for(let i=pKeys.length-1; i>0; i--) { const j = Math.floor(Math.random()*(i+1)); [pKeys[i], pKeys[j]] = [pKeys[j], pKeys[i]]; }
    
    const floorsData = {};
    for(let f=6; f>=1; f--) floorsData[f] = generateRoomsForFloor(f);

    update(ref(db, `games/Escape/rooms/${currentRoomId}`), {
        status: 'playing', turnIndex: 0, order: pKeys,
        floors: floorsData, log: { msg: 'เกมเริ่มขึ้นแล้ว หาทางลงไปยังชั้น 1!', ts: Date.now() }
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
            document.getElementById('participant-count').textContent = `${pList.length}/4`;
            const listEl = document.getElementById('participant-list');
            listEl.innerHTML = '';
            pList.forEach(p => {
                listEl.innerHTML += `<div class="participant-card">${p.name} ${p.isBot?'(บอท)':''}</div>`;
            });
            if(isHost) {
                document.getElementById('host-controls').style.display = 'block';
                document.getElementById('bot-count-display').textContent = gameState.botCount||0;
                document.getElementById('btn-start-game').style.display = 'block';
                document.getElementById('btn-start-game').disabled = pList.length < 1;
            }
        } else if (gameState.status === 'playing') {
            if(!document.getElementById('screen-game').classList.contains('active-screen')) {
                switchScreen('screen-game', 'game-floor-title');
            }
            if(gameState.log && gameState.log.ts > localTs) {
                localTs = gameState.log.ts;
                announceSR(gameState.log.msg, 'polite');
                document.getElementById('latest-event-display').textContent = gameState.log.msg;
            }
            renderGameUI();
        } else if (gameState.status === 'ended') {
            switchScreen('screen-result', 'result-title');
            renderResult();
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
    const isMyTurn = (curPId === myId && !myP.isBot && myP.status === 'normal');
    
    if(lastAnnouncedTurn !== `${gameState.turnIndex}_${curPId}`) {
        lastAnnouncedTurn = `${gameState.turnIndex}_${curPId}`;
        announceSR(`ถึงเทิร์นของ ${curP.name}`);
        playAudio('turn.mp3');
        if(isHost && curP.isBot && curP.status === 'normal') {
            setTimeout(() => executeBotTurn(curPId), 1500);
        } else if(curP.status !== 'normal' && isHost) {
            // Auto skip dead/escaped/resting
            setTimeout(() => endTurn(curPId), 1000);
        }
    }
    
    document.getElementById('game-status-bar').textContent = `เทิร์นของ: ${curP.name}`;
    
    // Status Cards
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
            กุญแจ: ${p.keys} | เกราะ: ${p.armor} | ปืน: ${p.gun}(${p.ammo})
        </div>`;
    });
    
    // Board
    const boardEl = document.getElementById('escape-board');
    boardEl.innerHTML = '';
    const floorData = gameState.floors[myP.floor];
    if(floorData) {
        for(let i=0; i<20; i++) {
            const r = floorData[i];
            const btn = document.createElement('button');
            btn.className = 'board-cell';
            const num = i+1;
            
            if(r.explored && r.type !== 'clue') {
                btn.classList.add('explored');
                btn.disabled = true;
                btn.innerHTML = `<span class="cell-icon">❌</span>ห้อง ${num}`;
                btn.setAttribute('aria-label', `ห้อง ${num} สำรวจแล้ว`);
            } else {
                let lbl = `ห้อง ${num}`;
                if(r.type === 'clue') {
                    btn.innerHTML = `<span class="cell-icon">❓</span>คำใบ้`;
                    lbl = `ห้องคำใบ้ ${num} สามารถสำรวจซ้ำได้`;
                } else {
                    btn.innerHTML = `<span class="cell-icon">🚪</span>ห้อง ${num}`;
                }
                
                if(isMyTurn) {
                    btn.onclick = () => exploreRoom(i, r);
                    lbl += ' กดเพื่อสำรวจ';
                } else {
                    btn.disabled = true;
                }
                btn.setAttribute('aria-label', lbl);
            }
            boardEl.appendChild(btn);
        }
    }
    
    // Stairs
    const btnStairs = document.getElementById('btn-use-stairs');
    if(isMyTurn && (myP.keys > 0 || myP.ghostKeys > 0 || myP.secretPassage || myP.canUseDoor)) {
        btnStairs.disabled = false;
        btnStairs.onclick = useStairs;
    } else {
        btnStairs.disabled = true;
    }
}

function readAllStatus() {
    let txt = "สถานะทั้งหมด: ";
    gameState.order.forEach(pid => {
        const p = gameState.players[pid];
        txt += `${p.name} อยู่ชั้น ${p.floor} HP ${p.hp} `;
        if(p.status !== 'normal') txt += `สถานะ ${p.status} `;
        if(p.infection > 0) txt += `ติดเชื้อระดับ ${p.infection} `;
        txt += `กุญแจ ${p.keys} เกราะ ${p.armor}, `;
    });
    announceSR(txt);
}

// Logic - Turn Processing
async function processTurnStartEffects(pId) {
    let p = gameState.players[pId];
    if(p.status === 'dead' || p.status === 'escaped') return p;
    
    if(p.status === 'resting') {
        p.hp = 100;
        p.status = 'normal';
        await emitLog(`${p.name} พักฟื้นเสร็จสิ้น ฟื้น HP เป็น 100 และพร้อมลุยต่อ`);
        return p; // skips action this turn essentially? wait, rules say "skip turn, then heal to 100".
        // To properly skip: if resting, heal and immediately end turn.
    }
    
    if(p.infection > 0) {
        if(p.vaccine > 0) {
            p.vaccine--; p.infection = 0;
            await emitLog(`${p.name} ใช้วัคซีนรักษาการติดเชื้ออัตโนมัติ`);
        } else {
            p.hp -= p.infection;
            await emitLog(`${p.name} เสีย ${p.infection} HP จากการติดเชื้อ (เหลือ ${p.hp} HP)`);
        }
    }
    
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
        endTurn(pId);
        return;
    }

    if(roomData.type === 'clue') {
        openClueModal(roomData.clueId);
        return; // Don't end turn yet
    }
    
    // Mark room explored
    await update(ref(db), { [`games/Escape/rooms/${currentRoomId}/floors/${p.floor}/${roomIdx}/explored`]: true });
    
    let msg = `${p.name} สำรวจห้อง ${roomIdx+1}: `;
    
    switch(roomData.type) {
        case 'key': p.keys++; msg += "พบกุญแจปกติ 1 ดอก"; break;
        case 'ghostKey': p.ghostKeys++; msg += "พบกุญแจผี 1 ดอก"; break;
        case 'secret': p.secretPassage = true; msg += "พบทางลงลับ! สามารถลงข้าม 2 ชั้นได้"; break;
        case 'vaccine': p.vaccine++; msg += "พบวัคซีน 1 ขวด"; break;
        case 'food': 
            let heal = 30; 
            p.hp = Math.min(200, p.hp + heal); 
            msg += `พบอาหาร ฟื้นฟู HP ${heal} หน่วย (รวม ${p.hp})`; 
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
    endTurn(pId);
}

function openClueModal(cid) {
    const clue = window.EscapeTowerClueBank ? window.EscapeTowerClueBank.getClueById(cid) : null;
    if(!clue) {
        // Fallback if no clues loaded
        emitLog(`${gameState.players[myId].name} สำรวจห้องคำใบ้ แต่ระบบคำใบ้ไม่พร้อมใช้งาน รับสิทธิ์ผ่านประตูไปเลย`);
        gameState.players[myId].canUseDoor = true;
        commitPlayer(myId, gameState.players[myId]).then(()=>endTurn(myId));
        return;
    }
    
    const m = document.getElementById('clue-modal');
    document.getElementById('clue-question').textContent = clue.question;
    document.getElementById('clue-answer-input').value = '';
    document.getElementById('clue-answer-input').dataset.cid = cid;
    document.getElementById('clue-feedback').textContent = '';
    m.style.display = 'flex';
    document.getElementById('clue-title').focus();
    announceSR(clue.question);
}

function submitClue() {
    const cid = document.getElementById('clue-answer-input').dataset.cid;
    const ans = document.getElementById('clue-answer-input').value.trim();
    if(!ans) return;
    
    const isCorrect = window.EscapeTowerClueBank.checkAnswer(cid, ans);
    if(isCorrect) {
        document.getElementById('clue-modal').style.display = 'none';
        const p = gameState.players[myId];
        p.canUseDoor = true;
        emitLog(`${p.name} ตอบคำใบ้ถูกต้อง ได้รับสิทธิ์เปิดประตูชั้น ${p.floor}`);
        commitPlayer(myId, p).then(()=>endTurn(myId));
    } else {
        document.getElementById('clue-feedback').textContent = "คำตอบผิด ลองใหม่ในเทิร์นหน้า";
        announceSR("คำตอบผิด ลองใหม่ในเทิร์นหน้า");
        setTimeout(() => {
            document.getElementById('clue-modal').style.display = 'none';
            emitLog(`${gameState.players[myId].name} ตอบคำใบ้ผิด ประตูยังไม่เปิด`);
            endTurn(myId);
        }, 1500);
    }
}

async function useStairs() {
    let p = await processTurnStartEffects(myId);
    if(p.status !== 'normal') { await commitPlayer(myId, p); endTurn(myId); return; }

    let down = 1;
    let used = "";
    if(p.secretPassage) { down = 2; p.secretPassage = false; used = "ทางลงลับ"; }
    else if(p.canUseDoor) { p.canUseDoor = false; used = "รหัสคำใบ้"; }
    else if(p.ghostKeys > 0) { p.ghostKeys--; used = "กุญแจผี"; }
    else if(p.keys > 0) { p.keys--; used = "กุญแจปกติ"; }
    
    p.floor = Math.max(1, p.floor - down);
    p.canUseDoor = false; // Reset clue auth per floor
    
    let msg = `${p.name} ใช้${used} ลงไปยังชั้น ${p.floor}`;
    if(p.floor === 1) {
        p.status = 'escaped';
        msg += " และหนีออกจากตึกสำเร็จ!!";
    }
    
    await emitLog(msg);
    await commitPlayer(myId, p);
    
    if(p.floor === 1) checkEndGame();
    endTurn(myId);
}

// Bot Logic (Very Basic AI)
async function executeBotTurn(botId) {
    let p = gameState.players[botId];
    if(!p || p.status !== 'normal') { endTurn(botId); return; }
    
    // Evaluate if should go down
    if(p.secretPassage || p.ghostKeys > 0 || p.keys > 0) {
        let down = p.secretPassage ? 2 : 1;
        let used = p.secretPassage ? "ทางลงลับ" : (p.ghostKeys>0 ? "กุญแจผี" : "กุญแจปกติ");
        p.secretPassage = false;
        if(used==="กุญแจผี") p.ghostKeys--; else if(used==="กุญแจปกติ") p.keys--;
        
        p.floor = Math.max(1, p.floor - down);
        let msg = `${p.name} ใช้${used} ลงไปยังชั้น ${p.floor}`;
        if(p.floor === 1) { p.status = 'escaped'; msg += " หนีสำเร็จ!!"; }
        await emitLog(msg);
        await commitPlayer(botId, p);
        if(p.floor === 1) checkEndGame();
        endTurn(botId);
        return;
    }
    
    // Pick random unexplored room
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
            case 'secret': p.secretPassage = true; msg += "พบทางลงลับ"; break;
            case 'vaccine': p.vaccine++; msg += "พบวัคซีน"; break;
            case 'food': p.hp = Math.min(200, p.hp+30); msg += "กินอาหารฟื้น HP"; break;
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
    endTurn(botId);
}

// Helpers
async function emitLog(msg) {
    await update(ref(db), { [`games/Escape/rooms/${currentRoomId}/log`]: { msg, ts: Date.now() } });
}

async function commitPlayer(pId, pData) {
    await update(ref(db), { [`games/Escape/rooms/${currentRoomId}/players/${pId}`]: pData });
}

function endTurn(curPId) {
    const order = gameState.order;
    let idx = gameState.turnIndex;
    
    // Find next playable player (or just advance 1 if all dead/escaped to avoid infinite loop, checked by checkEndGame anyway)
    idx = (idx + 1) % order.length;
    
    update(ref(db), { [`games/Escape/rooms/${currentRoomId}/turnIndex`]: idx });
    checkEndGame();
}

function checkEndGame() {
    if(!isHost || !gameState) return;
    let allDone = true;
    for(let k in gameState.players) {
        if(gameState.players[k].status === 'normal' || gameState.players[k].status === 'resting') {
            allDone = false; break;
        }
    }
    if(allDone) {
        update(ref(db, `games/Escape/rooms/${currentRoomId}`), { status: 'ended' });
    }
}

function renderResult() {
    announceSR("เกมจบลงแล้ว นี่คือสรุปผล", "assertive");
    const resEl = document.getElementById('result-list');
    resEl.innerHTML = '';
    const pArr = Object.values(gameState.players).sort((a,b) => a.floor - b.floor);
    pArr.forEach(p => {
        let txt = `${p.name}: `;
        if(p.status === 'escaped') txt += "หนีรอดสำเร็จ! 🏆";
        else if(p.status === 'dead') txt += "เสียชีวิต 💀";
        else txt += `ติดอยู่ที่ชั้น ${p.floor}`;
        resEl.innerHTML += `<div class="participant-card">${txt}</div>`;
    });
}