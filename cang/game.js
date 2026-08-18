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
// Audio Setup (คงไว้ 100%)
// =========================================================
const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
const audioBuffers = {};
let bgmSource = null;

const soundFiles = {
    '1': 'audio/1.mp3', 'select': 'audio/select.mp3', 'start': 'audio/start.mp3',
    'bgm': 'audio/bgm.mp3', 'jua': 'audio/jua.mp3', 'turn': 'audio/turn.mp3',
    'follow': 'audio/follow.mp3', 'cang': 'audio/cang.mp3', 'cang25': 'audio/cang25.mp3',
    'win': 'audio/win.mp3', 'lost': 'audio/lost.mp3', '60': 'audio/60.mp3', 'no': 'audio/no.mp3'
};

async function loadSounds() {
    for (let key in soundFiles) {
        try {
            const response = await fetch(soundFiles[key]);
            const arrayBuffer = await response.arrayBuffer();
            const audioBuffer = await audioCtx.decodeAudioData(arrayBuffer);
            audioBuffers[key] = audioBuffer;
        } catch (e) { console.warn("Could not load sound", key, e); }
    }
}
loadSounds();

function playSound(key, loop = false) {
    if(audioCtx.state === 'suspended') audioCtx.resume();
    if(!audioBuffers[key]) return null;
    const source = audioCtx.createBufferSource();
    source.buffer = audioBuffers[key];
    source.connect(audioCtx.destination);
    source.loop = loop;
    source.start(0);
    return source;
}

function stopBGM() {
    if (bgmSource) { try { bgmSource.stop(); } catch(e) {} bgmSource = null; }
}

function playBGM() { stopBGM(); bgmSource = playSound('bgm', true); }

document.body.addEventListener('click', () => { if(audioCtx.state === 'suspended') audioCtx.resume(); }, { once: false });

function triggerSound(soundKey) {
    if(myPeerId) playSound(soundKey);
    guestConnections.forEach(c => c.send({ type: 'PLAY_SOUND', sound: soundKey }));
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

let peer = null, myPeerId = null, roomId = null, isHost = false;
let hostConnection = null, guestConnections = []; 
let roomPlayers = [], botColors = [], currentBotCount = 0, selectedColorId = 'red';
let deck = [], discardPile = [];
let gameState = { turnIndex: 0, players: [], status: 'WAITING', topCardOwnerId: null, flowSourceId: null, skipPreVotes: [] };
let localPlayerState = { hand: [], hasDrawn: false, hasDiscarded: false, discardedRank: null, points: 300, turnCount: 0, hasFinishedFirstTurn: false, hasFlowedThisTurn: false };
let globalPlayersMap = {}; 
let turnTimerInterval = null;
let announcerTimeout = null;
let hostHeartbeatTimer = null;
let heartbeatInterval = null;

function triggerHostDisconnect() {
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

let currentTopCardText = "ยังไม่มีไพ่บนกอง";

// Shortcut Keys
document.addEventListener('keydown', (e) => {
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
    const setupDisp = window.getComputedStyle(document.getElementById('setup-screen')).display;
    if (setupDisp !== 'none') homeBtn.style.display = 'block';
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
    if(!p) return "ผู้เล่น";
    if(id === myPeerId) return `คุณสี${p.colorName}`;
    if(p.isBot) return `บอทสี${p.colorName}`;
    return `เพื่อนสี${p.colorName}`;
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
// Network & Lobby Logic
// =========================================================
window.onload = () => {
    myPeerId = "cang-client-" + Math.random().toString(36).substring(2, 8);
    peer = new Peer(myPeerId);
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
                roomPlayers = [{ id: myPeerId, color: selectedColorId }];
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
    isHost = false; roomId = targetRoomId;
    announce(`กำลังเข้าร่วมห้อง ${roomId}`, false);
    document.getElementById('setup-screen').style.display = 'none';
    document.getElementById('waiting-room-screen').style.display = 'flex';
    document.getElementById('host-bot-section').style.display = 'none';
    updateHomeBtnVisibility();
    
    hostConnection = peer.connect("cang-host-" + roomId);
    hostConnection.on('open', () => { 
        hostConnection.send({ type: 'JOIN_ROOM' }); 
        announce(`เข้าร่วมห้องสำเร็จ รอแจกไพ่`, false); 
        resetHeartbeat();
    });
    hostConnection.on('data', (data) => {
        resetHeartbeat();
        if (data.type === 'HEARTBEAT') return;
        if (data.type === 'LOBBY_UPDATE') { 
            roomPlayers = data.roomPlayers; currentBotCount = data.botCount; botColors = data.botColors; 
            if (data.joinOrder) window.joinOrder = data.joinOrder; 
            updateLobbyUI(); 
            if (!isHost && !window.hasAnnouncedMyColor) {
                let myP = roomPlayers.find(p => p.id === myPeerId);
                if (myP) {
                    window.hasAnnouncedMyColor = true;
                    let c = COLORS.find(x => x.id === myP.color);
                    if (c) announce(`คุณได้รับสี${c.name}`, false);
                }
            }
        }
        if (data.type === 'START_COUNTDOWN') { 
            if (!window.isStartingRound) { window.isStartingRound = true; doCountdownAndStart(); }
        }
        if (data.type === 'SYNC_STATE') renderClientGame(data.publicState, data.privateState);
        if (data.type === 'ANIM_SC_FLOAT') showFloatingScore(data.targetId, data.amt);
        if (data.type === 'SHOW_RESULT') showResultModal(data.title, data.detail, data.isEnd, data.winners, data.losers);
        if (data.type === 'ANNOUNCE') announce(data.msg);
        if (data.type === 'TIMER_UPDATE') window.updateTimerUI(data.timeLeft);
        if (data.type === 'PLAY_SOUND') playSound(data.sound);
        if (data.type === 'SHOW_TRANSITION') window.showTransition(data.msg, data.duration);
        if (data.type === 'HIDE_TRANSITION') window.hideTransition();
        if (data.type === 'KANG_ANIMATION') window.playKangAnimation(data.callerId, data.isKang25);
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
                
                roomPlayers.push({ id: conn.peer, color: randomColor });
                window.joinOrder.push(conn.peer); 
                guestConnections.push(conn); 
                broadcastLobby();
                broadcastAnnounce(`มีผู้เล่นเข้าร่วมห้องใหม่ รวม ${roomPlayers.length + currentBotCount} ตัว`);
                triggerSound('select');
            }
            if (data.type === 'CHANGE_COLOR') {
                let p = roomPlayers.find(x => x.id === data.playerId);
                let taken = roomPlayers.map(x=>x.color).concat(botColors);
                if(p && !taken.includes(data.colorId)) { p.color = data.colorId; broadcastLobby(); triggerSound('select'); }
            }
            if (data.type === 'PLAYER_ACTION') processPlayerAction(data.playerId, data.action, data.cardIndex);
        });
    });
}

function broadcastLobby() {
    let data = { type: 'LOBBY_UPDATE', roomPlayers, botCount: currentBotCount, botColors, joinOrder: window.joinOrder };
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
    } else if (delta < 0 && currentBotCount > 0) {
        currentBotCount--; botColors.pop(); 
        for (let i = window.joinOrder.length - 1; i >= 0; i--) {
            if (window.joinOrder[i].startsWith('bot_')) { window.joinOrder.splice(i, 1); break; }
        }
    }
    broadcastLobby(); triggerSound('select');
};

function updateLobbyUI() {
    document.getElementById('display-room-id').innerText = roomId;
    document.getElementById('bot-count-text').innerText = `${currentBotCount} ตัว`;
    
    globalPlayersMap = {};
    roomPlayers.forEach(p => { let c = COLORS.find(x => x.id === p.color); globalPlayersMap[p.id] = { isBot: false, colorName: c ? c.name : '?', colorHex: c ? c.hex : '#ffffff' }; });
    botColors.forEach((b, i) => { let c = COLORS.find(x => x.id === b); globalPlayersMap['bot_'+i] = { isBot: true, colorName: c ? c.name : '?', colorHex: c ? c.hex : '#ffffff' }; });

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

    let taken = roomPlayers.map(p=>p.color).concat(botColors);
    let myP = roomPlayers.find(p=>p.id===myPeerId);
    if (myP && taken.filter(c => c === myP.color).length > 1) {
        let avail = COLORS.find(c => !taken.includes(c.id));
        if(avail) { selectedColorId = avail.id; myP.color = selectedColorId; if(isHost) { broadcastLobby(); triggerSound('select'); } }
    }

    const colDiv = document.getElementById('color-selection'); colDiv.innerHTML = '';
    COLORS.forEach(c => {
        let isTakenByOther = taken.includes(c.id) && c.id !== selectedColorId;
        let label = document.createElement('label');
        label.className = `radio-color-label ${c.id === selectedColorId ? 'selected':''} ${isTakenByOther ? 'disabled-color':''}`;
        label.style.backgroundColor = c.hex;
        
        let input = document.createElement('input'); 
        input.type = 'radio'; input.name = 'player-color'; input.value = c.id;
        input.setAttribute('aria-label', `เลือกสี${c.name}`);
        if(c.id === selectedColorId) input.checked = true; 
        if(isTakenByOther) input.disabled = true;
        
        input.onchange = () => {
            if(!isTakenByOther) {
                selectedColorId = c.id; announce(`คุณเลือกเปลี่ยนสีเป็นสี${c.name}`, false);
                if(isHost) { 
                    let p = roomPlayers.find(x=>x.id===myPeerId); 
                    if(p) { p.color = c.id; broadcastLobby(); triggerSound('select'); }
                } else { hostConnection.send({ type: 'CHANGE_COLOR', playerId: myPeerId, colorId: c.id }); }
            }
        };
        
        let span = document.createElement('span'); span.innerText = c.name;
        label.appendChild(input); label.appendChild(span); colDiv.appendChild(label);
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
    guestConnections.forEach(c => c.send({ type: 'START_COUNTDOWN' }));
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
                            let botIdx = parseInt(id.split('_')[1]); let bColor = botColors[botIdx];
                            if (bColor) gameState.players.push({ id: id, name: 'บอท '+COLORS.find(c=>c.id===bColor).name, colorInfo: COLORS.find(c=>c.id===bColor), isBot: true, points: 300, hand: [], isOut: false, turnCount: 0, hasFinishedFirstTurn: false, hasFlowedThisTurn: false });
                        } else {
                            let p = roomPlayers.find(x => x.id === id);
                            if (p) gameState.players.push({ id: p.id, name: p.id===myPeerId?'Host':'Player', colorInfo: COLORS.find(c=>c.id===p.color), isBot: false, points: 300, hand: [], isOut: false, turnCount: 0, hasFinishedFirstTurn: false, hasFlowedThisTurn: false });
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
    
    gameState.players.forEach((p) => {
        if(p.points > 0) { 
            p.hand = deck.splice(0, 5); p.isOut = false; p.hasDrawnTurn = false; p.hasDiscardedTurn = false; p.hasFlowedThisTurn = false; p.discardedRankThisTurn = null; p.turnCount = 0; p.hasFinishedFirstTurn = false;
        } else { p.isOut = true; } 
    });
    
    syncStateToAll(); startPreGameTimer(); broadcastAnnounce("เปิดรอบใหม่! มีใครจะแคง 25 ไหม?");
}

function startPreGameTimer() {
    clearInterval(turnTimerInterval);
    let timeLeft = 40; broadcastTimer(timeLeft);
    turnTimerInterval = setInterval(() => {
        timeLeft--; broadcastTimer(timeLeft);
        if(timeLeft <= 0) { 
            clearInterval(turnTimerInterval); 
            if (gameState.status === 'PRE_GAME') {
                triggerSound('60'); gameState.status = 'PLAYING';
                broadcastAnnounce("📢 หมดเวลาดูไพ่ 40 วินาที ไม่มีใครกดแคง! ขอเชิญผู้เล่นแรกจั่วไพ่ใบแรกและทิ้งลงกองเพื่อเริ่มเกมได้เลยครับ");
                syncStateToAll(); startTurnTimer(); checkBotTurn();
            }
        }
    }, 1000);
}

function startTurnTimer() {
    clearInterval(turnTimerInterval);
    let timeLeft = 40; broadcastTimer(timeLeft);
    turnTimerInterval = setInterval(() => {
        timeLeft--; broadcastTimer(timeLeft);
        if(timeLeft <= 0) { clearInterval(turnTimerInterval); handleTurnTimeout(); }
    }, 1000);
}

function broadcastTimer(t) {
    if(myPeerId) window.updateTimerUI(t);
    guestConnections.forEach(c => c.send({ type: 'TIMER_UPDATE', timeLeft: t }));
}

window.updateTimerUI = (timeLeft) => {
    const timerBar = document.getElementById('timer-bar-fill');
    if(timerBar) {
        let maxTime = 40; let pct = Math.max(0, (timeLeft / maxTime) * 100);
        timerBar.style.width = pct + '%'; timerBar.style.background = timeLeft <= 10 ? 'var(--red)' : 'var(--orange)';
    }
};

function handleTurnTimeout() {
    let cp = gameState.players[gameState.turnIndex];
    if(cp.isBot || gameState.status !== 'PLAYING') return;
    
    if (discardPile.length === 0 && !cp.hasDiscardedTurn) {
        if (!cp.hasDrawnTurn) processPlayerAction(cp.id, 'DRAW');
        if (cp.hand.length > 0) processPlayerAction(cp.id, 'DISCARD', 0);
    } else {
        if (!cp.hasDrawnTurn) processPlayerAction(cp.id, 'DRAW');
        if (!cp.hasDiscardedTurn && cp.hand.length > 0) processPlayerAction(cp.id, 'DISCARD', 0);
    }
    if (gameState.status === 'PLAYING') processPlayerAction(cp.id, 'END_TURN');
}

function syncStateToAll() {
    let activePlayers = gameState.players.filter(p => !p.isOut);
    if(activePlayers.length <= 1) return;

    let publicState = {
        status: gameState.status, turnId: gameState.players[gameState.turnIndex].id, topDiscard: discardPile[discardPile.length - 1],
        deckCount: deck.length, discardPileCount: discardPile.length, skipVotes: gameState.skipPreVotes || [],
        playersInfo: gameState.players.map(p => ({
            id: p.id, colorHex: p.colorInfo.hex, colorName: p.colorInfo.name, isBot: p.isBot, points: p.points, cardCount: p.hand.length, isOut: p.isOut
        }))
    };

    gameState.players.forEach(p => {
        if(!p.isBot) {
            let privateState = { hand: p.hand, hasDrawn: p.hasDrawnTurn, hasDiscarded: p.hasDiscardedTurn, discardedRank: p.discardedRankThisTurn, points: p.points, turnCount: p.turnCount, hasFinishedFirstTurn: p.hasFinishedFirstTurn, hasFlowedThisTurn: p.hasFlowedThisTurn };
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
// [พื้นที่สำหรับพัฒนาต่อ #2] ระบบ Animation และ Visual Layer
// =========================================================

// สร้าง Particle สวยงามแบบไม่ต้องพึ่งพา Library เพิ่มเติม
window.createParticles = (type, x, y) => {
    const container = document.createElement('div');
    container.style.position = 'fixed'; container.style.left = x + 'px'; container.style.top = y + 'px';
    container.style.pointerEvents = 'none'; container.style.zIndex = '9999';
    container.setAttribute('aria-hidden', 'true');
    document.body.appendChild(container);

    let colors = ['#ffd700', '#ffea00', '#fff'];
    if (type === 'coin') colors = ['#00ff00', '#a8ff78'];
    if (type === 'impact') colors = ['#ff4d4d', '#ff7675'];

    for(let i=0; i < (type === 'gold' ? 15 : 8); i++) {
        let p = document.createElement('div');
        p.className = `particle p-${type}`;
        p.style.backgroundColor = colors[Math.floor(Math.random() * colors.length)];
        
        let angle = Math.random() * Math.PI * 2;
        let dist = Math.random() * 60 + 20;
        p.style.setProperty('--dx', Math.cos(angle)*dist + 'px');
        p.style.setProperty('--dy', Math.sin(angle)*dist + 'px');
        
        container.appendChild(p);
    }
    setTimeout(() => container.remove(), 1200);
}

// Floating Score Animation (เหรียญเด้ง)
window.showFloatingScore = (targetId, amt) => {
    let isPositive = amt > 0;
    let text = (isPositive ? '+' : '') + amt;
    let color = isPositive ? 'var(--green)' : 'var(--red)';

    let avatar = document.getElementById(`avatar-players-status-bar-${targetId}`);
    if(!avatar) avatar = document.getElementById(`avatar-result-status-bar-${targetId}`);

    let x = window.innerWidth / 2; let y = window.innerHeight / 2;
    if(avatar) {
        let rect = avatar.getBoundingClientRect();
        x = rect.left + rect.width / 2; y = rect.top;
    }

    let floatEl = document.createElement('div');
    floatEl.className = `floating-score ${isPositive ? 'score-gain' : 'score-loss'}`;
    floatEl.innerText = text;
    floatEl.style.left = x + 'px'; floatEl.style.top = (y - 20) + 'px';
    floatEl.style.color = color;
    floatEl.setAttribute('aria-hidden', 'true'); 
    document.body.appendChild(floatEl);

    window.createParticles(isPositive ? 'coin' : 'impact', x, y);
    setTimeout(() => floatEl.remove(), 2000);
};

// KANG Boss Moment Cinematic
window.playKangAnimation = (callerId, isKang25) => {
    const animScreen = document.getElementById('kang-animation-screen');
    const text1 = document.getElementById('kang-anim-text1');
    const text2 = document.getElementById('kang-anim-text2');
    const avatar = document.getElementById('kang-anim-avatar');

    animScreen.style.display = 'flex';
    text1.style.opacity = '0';
    text2.style.opacity = '0';
    
    let pInfo = globalPlayersMap[callerId];
    let pColorHex = pInfo ? pInfo.colorHex : 'white';
    
    avatar.style.color = pColorHex;
    avatar.className = 'kang-anim-active';

    playSound(isKang25 ? 'cang25' : 'cang');

    let callerName = resolveName(callerId);
    let word = isKang25 ? 'แคง 25' : 'แคง';

    setTimeout(() => {
        let t1 = `💥 ${callerName} ประกาศ ${word}!`;
        text1.innerText = t1; text1.style.opacity = '1'; announce(t1, false);
        window.createParticles('gold', window.innerWidth/2, window.innerHeight/2);
    }, 1200);

    setTimeout(() => {
        let t2 = `เปิดไพ่เพื่อวัดดวงกัน!`;
        text2.innerText = t2; text2.style.opacity = '1'; announce(t2, false);
    }, 3000);

    setTimeout(() => {
        animScreen.style.display = 'none'; avatar.className = 'kang-avatar-epic';
    }, 4500);
};

function broadcastKangAnimation(callerId, isKang25) {
    gameState.status = 'TRANSITION'; clearInterval(turnTimerInterval); syncStateToAll();
    let data = { type: 'KANG_ANIMATION', callerId, isKang25 };
    if(myPeerId) window.playKangAnimation(callerId, isKang25);
    guestConnections.forEach(c => c.send(data));
}

// Client UI Renderer
function renderClientGame(publicState, privateState) {
    // 1. จัดการ Status Bar ผู้เล่น
    const statusGroup = document.getElementById('players-status-bar');
    statusGroup.innerHTML = '';
    
    publicState.playersInfo.forEach((p, idx) => {
        if(p.isOut) return;
        let pDiv = document.createElement('div');
        pDiv.className = 'player-status-card glass-panel ' + (publicState.turnId === p.id && publicState.status === 'PLAYING' ? 'active-turn' : '');
        pDiv.id = `players-status-bar-${p.id}`;
        
        let pAvatar = document.createElement('div');
        pAvatar.className = 'player-avatar';
        pAvatar.id = `avatar-players-status-bar-${p.id}`;
        pAvatar.style.color = p.colorHex;
        pAvatar.innerHTML = '👤';
        pAvatar.setAttribute('aria-hidden', 'true');

        let nameEl = document.createElement('div'); nameEl.className = 'player-name'; nameEl.innerText = resolveName(p.id);
        let coinEl = document.createElement('div'); coinEl.className = 'player-coins'; coinEl.innerText = p.points + ' 💰';
        let cCount = document.createElement('div'); cCount.className = 'player-card-count'; cCount.innerText = '🎴 ' + p.cardCount;
        
        let contentDiv = document.createElement('div');
        contentDiv.appendChild(nameEl); contentDiv.appendChild(coinEl); contentDiv.appendChild(cCount);
        
        pDiv.appendChild(pAvatar);
        pDiv.appendChild(contentDiv);
        statusGroup.appendChild(pDiv);
    });

    let topC = publicState.topDiscard;
    currentTopCardText = topC ? `ไพ่บนกองคือ ${THAI_RANKS[topC.rank]} ${THAI_SUITS[topC.suit]}` : 'ยังไม่มีไพ่บนกองทิ้ง';

    // 2. จัดการ Turn Indicator 
    const turnEl = document.getElementById('turn-indicator');
    if (publicState.status === 'PRE_GAME') {
        turnEl.innerText = "ช่วงดูไพ่: รอคนแคง (40 วิ)";
        turnEl.className = "turn-indicator-text highlight-pregame";
    } else {
        if(publicState.turnId === myPeerId) { turnEl.innerText = "🔥 ตาของคุณแล้ว!"; turnEl.className = "turn-indicator-text my-turn-glow"; }
        else { turnEl.innerText = `ตาของ ${resolveName(publicState.turnId)}`; turnEl.className = "turn-indicator-text"; }
    }

    // 3. จัดการ Action Buttons
    let isMyTurn = (publicState.turnId === myPeerId);
    let btnDraw = document.getElementById('btn-draw'); let btnEnd = document.getElementById('btn-end-turn');
    let btnKang = document.getElementById('btn-kang'); let btnKang25 = document.getElementById('btn-kang25');
    let btnSkip = document.getElementById('btn-skip-pre'); let btnSpecial = document.getElementById('btn-special');

    if (publicState.status === 'PRE_GAME') {
        btnDraw.style.display = 'none'; btnEnd.style.display = 'none';
        btnKang.style.display = 'inline-block'; btnKang25.style.display = 'inline-block'; btnSkip.style.display = 'inline-block';
        
        let isFirstPlayer = (publicState.playersInfo[0].id === myPeerId);
        let handSum = privateState.hand.reduce((s, c) => s + c.val, 0);
        let hasKang25Cond = (handSum >= 25);
        let hasVoted = publicState.skipVotes.includes(myPeerId);
        
        btnKang.disabled = !isFirstPlayer || hasKang25Cond;
        btnKang25.disabled = !hasKang25Cond;
        btnSkip.disabled = hasVoted;
        
    } else if (publicState.status === 'PLAYING') {
        btnKang25.style.display = 'none'; btnSkip.style.display = 'none';
        btnDraw.style.display = 'inline-block'; btnEnd.style.display = 'inline-block'; btnKang.style.display = 'inline-block';
        
        btnDraw.disabled = !(isMyTurn && !privateState.hasDrawn && !privateState.hasFlowedThisTurn && publicState.deckCount > 0);
        btnEnd.disabled = !(isMyTurn && (privateState.hasDiscarded || privateState.hasFlowedThisTurn));
        
        let canKang = isMyTurn && !privateState.hasDrawn && !privateState.hasFlowedThisTurn && publicState.discardPileCount > 0;
        btnKang.disabled = !canKang;
    }

    // Special Win Rule Check
    let swType = getSpecialWinType(privateState.hand, publicState.playersInfo.find(p=>p.id===myPeerId));
    if (swType && (publicState.status === 'PRE_GAME' || publicState.status === 'PLAYING')) {
        currentSpecialWinType = swType;
        btnSpecial.innerText = `ชนะพิเศษ: ${swType}!`;
        btnSpecial.style.display = 'inline-block';
    } else {
        currentSpecialWinType = null;
        btnSpecial.style.display = 'none';
    }

    // 4. แสดงผลจำนวนไพ่กองจั่ว (Visual Real Card Back)
    document.getElementById('deck-count-text').innerText = publicState.deckCount;
    if(publicState.deckCount > 0) {
        document.getElementById('deck-group-box').style.opacity = '1';
    } else {
        document.getElementById('deck-group-box').style.opacity = '0.5';
    }

    // 5. แสดงผลกองทิ้ง (Visual Real Card)
    const discardContainer = document.getElementById('discard-pile');
    discardContainer.innerHTML = '';
    if (topC) {
        let cardDiv = document.createElement('div');
        cardDiv.className = 'playing-card card-on-table anim-discard'; 
        let isRed = (topC.suit === '♥' || topC.suit === '♦');
        let suitColor = isRed ? 'var(--red)' : '#1a1a2e';
        
        cardDiv.innerHTML = `
            <div class="card-visual" style="color: ${suitColor};" aria-hidden="true">
                <div class="card-top">${topC.rank} <br> ${topC.suit}</div>
                <div class="card-center">${topC.suit}</div>
                <div class="card-bottom">${topC.rank} <br> ${topC.suit}</div>
            </div>
            <span class="sr-only">ไพ่บนกองทิ้งคือ ${THAI_RANKS[topC.rank]} ${THAI_SUITS[topC.suit]}</span>
        `;
        discardContainer.appendChild(cardDiv);
    } else {
        discardContainer.innerHTML = '<span class="sr-only">ยังไม่มีไพ่บนกองทิ้ง</span>';
    }

    // 6. แสดงผลไพ่บนมือ (Visual Real Card)
    const handContainer = document.getElementById('my-hand-ui');
    handContainer.innerHTML = '';
    privateState.hand.forEach((c, idx) => {
        let cardBtn = document.createElement('button');
        cardBtn.className = 'playing-card card-in-hand hand-card-btn';
        
        let isRed = (c.suit === '♥' || c.suit === '♦');
        let suitColor = isRed ? 'var(--red)' : '#1a1a2e'; 
        
        cardBtn.innerHTML = `
            <div class="card-visual" style="color: ${suitColor};" aria-hidden="true">
                <div class="card-top">${c.rank} <br> ${c.suit}</div>
                <div class="card-center" style="font-size:2rem;">${c.suit}</div>
                <div class="card-bottom">${c.rank} <br> ${c.suit}</div>
            </div>
            <span class="sr-only">ไพ่ใบที่ ${idx+1}: ${THAI_RANKS[c.rank]} ${THAI_SUITS[c.suit]}</span>
        `;
        
        let canDiscard = isMyTurn && privateState.hasDrawn && (!privateState.hasDiscarded || privateState.discardedRank === c.rank);
        let canFlow = false;
        
        if (isMyTurn && !privateState.hasDrawn && !privateState.hasDiscarded && publicState.status === 'PLAYING' && topC && c.rank === topC.rank) {
            canFlow = true;
        }

        if (canDiscard) {
            cardBtn.onclick = () => window.clientAction('DISCARD', idx);
            cardBtn.classList.add('can-discard-glow');
        } else if (canFlow) {
            cardBtn.onclick = () => window.clientAction('FLOW', idx);
            cardBtn.classList.add('can-flow-glow');
        } else {
            cardBtn.disabled = true;
        }
        
        handContainer.appendChild(cardBtn);
    });
}

window.clientAction = (action, cardIndex = null) => {
    if(isHost) processPlayerAction(myPeerId, action, cardIndex);
    else hostConnection.send({ type: 'PLAYER_ACTION', playerId: myPeerId, action, cardIndex });
};

// =========================================================
// Result Modal & End Ceremony
// =========================================================
window.showResultModal = (title, detail, isEnd, winners, losers) => {
    document.getElementById('result-title').innerText = title;
    document.getElementById('result-details').innerHTML = detail;
    
    let modal = document.getElementById('result-modal');
    modal.style.display = 'flex';
    
    // Ceremony Effect
    if (winners && winners.length > 0) {
        playSound('win');
        window.createParticles('gold', window.innerWidth/2, window.innerHeight/3);
    } else {
        playSound('lost');
    }
    
    if (isHost) {
        document.getElementById('btn-next-round').style.display = 'block';
        document.getElementById('btn-next-round').disabled = false;
        document.getElementById('guest-waiting-next-round').style.display = 'none';
    } else {
        document.getElementById('btn-next-round').style.display = 'none';
        document.getElementById('guest-waiting-next-round').style.display = 'block';
    }
    
    setTimeout(() => { document.getElementById('result-title').focus(); }, 100);
};

window.closeResultModal = () => {
    document.getElementById('result-modal').style.display = 'none';
    if(isHost) startNewRound();
};

// =========================================================
// Server Logic (Bot AI & Validation Placeholder)
// (ระบบ Logic เดิมทั้งหมดของ Server-side จะถูกรันต่อไปตามนี้)
// =========================================================
function processPlayerAction(playerId, action, cardIndex) {
    if(gameState.status === 'TRANSITION') return;

    let p = gameState.players.find(x => x.id === playerId);
    if(!p) return;

    if (action === 'DRAW') {
        if (deck.length > 0) {
            p.hand.push(deck.pop());
            p.hasDrawnTurn = true;
            triggerSound('jua');
            broadcastAnnounce(`[PID:${playerId}] จั่วไพ่ 1 ใบ`);
            syncStateToAll();
        }
    } 
    else if (action === 'DISCARD') {
        let discardedCard = p.hand.splice(cardIndex, 1)[0];
        discardPile.push(discardedCard);
        gameState.topCardOwnerId = playerId;
        p.hasDiscardedTurn = true;
        p.discardedRankThisTurn = discardedCard.rank;
        triggerSound('follow');
        broadcastAnnounce(`[PID:${playerId}] ทิ้งไพ่ ${THAI_RANKS[discardedCard.rank]} ${THAI_SUITS[discardedCard.suit]}`);
        syncStateToAll();
    }
    else if (action === 'FLOW') {
        let discardedCard = p.hand.splice(cardIndex, 1)[0];
        discardPile.push(discardedCard);
        
        let reward = (discardedCard.rank === 'A') ? 20 : 10;
        let sourceP = gameState.players.find(x => x.id === gameState.topCardOwnerId);
        
        p.hasFlowedThisTurn = true;
        triggerSound('follow');
        
        if (sourceP && sourceP.id !== p.id) {
            sourceP.points -= reward;
            p.points += reward;
            broadcastFloatSC(sourceP.id, -reward);
            broadcastFloatSC(p.id, reward);
        }
        gameState.topCardOwnerId = playerId;
        broadcastAnnounce(`[PID:${playerId}] ไหลไพ่ ${THAI_RANKS[discardedCard.rank]} ${THAI_SUITS[discardedCard.suit]} รับ ${reward} เหรียญ!`);
        syncStateToAll();
    }
    else if (action === 'END_TURN') {
        p.hasDrawnTurn = false;
        p.hasDiscardedTurn = false;
        p.hasFlowedThisTurn = false;
        p.discardedRankThisTurn = null;
        p.turnCount++;
        if (p.turnCount >= 1) p.hasFinishedFirstTurn = true;

        if (p.hand.length === 0) {
            handleRoundEndByCards(p); 
        } else {
            advanceTurn();
        }
    }
    else if (action === 'KANG' || action === 'KANG_25') {
        broadcastKangAnimation(playerId, action === 'KANG_25');
        setTimeout(() => { handleKangEnd(p, action === 'KANG_25'); }, 4600);
    }
    else if (action === 'SKIP_PRE') {
        if (!gameState.skipPreVotes.includes(playerId)) {
            gameState.skipPreVotes.push(playerId);
            let activeCount = gameState.players.filter(x=>!x.isOut).length;
            if (gameState.skipPreVotes.length >= activeCount) {
                gameState.status = 'PLAYING';
                broadcastAnnounce("ทุกคนพร้อมแล้ว เริ่มเกมได้!");
                startTurnTimer();
            }
            syncStateToAll();
        }
    }
    else if (action === 'SPECIAL_WIN') {
        // Logic ชนะพิเศษ แจกเหรียญ (คง logic เดิมไว้)
        broadcastKangAnimation(playerId, false);
        setTimeout(() => {
            let winAmt = 40;
            let winners = [playerId];
            let losers = [];
            gameState.players.forEach(op => {
                if(!op.isOut && op.id !== playerId) {
                    op.points -= winAmt;
                    p.points += winAmt;
                    losers.push(op.id);
                }
            });
            let det = `[PID:${playerId}] ชนะพิเศษ! รับคนละ ${winAmt} เหรียญ`;
            let isEnd = gameState.players.filter(x => x.points > 0).length <= 1;
            syncStateToAll();
            showResultModal("🏆 ชนะพิเศษ!", det, isEnd, winners, losers);
            let data = { type: 'SHOW_RESULT', title: "🏆 ชนะพิเศษ!", detail: det, isEnd: isEnd, winners, losers };
            guestConnections.forEach(c => c.send(data));
        }, 4600);
    }
}

function advanceTurn() {
    let activeP = gameState.players.filter(p => !p.isOut);
    let currIdx = activeP.findIndex(p => p.id === gameState.players[gameState.turnIndex].id);
    currIdx = (currIdx + 1) % activeP.length;
    gameState.turnIndex = gameState.players.findIndex(p => p.id === activeP[currIdx].id);
    
    syncStateToAll();
    startTurnTimer();
    checkBotTurn();
}

function checkBotTurn() {
    let cp = gameState.players[gameState.turnIndex];
    if (cp.isBot && gameState.status === 'PLAYING') {
        setTimeout(() => {
            let topC = discardPile[discardPile.length - 1];
            let matchIdx = cp.hand.findIndex(c => topC && c.rank === topC.rank);
            
            if (matchIdx !== -1) {
                processPlayerAction(cp.id, 'FLOW', matchIdx);
                setTimeout(() => processPlayerAction(cp.id, 'END_TURN'), 1500);
            } else {
                processPlayerAction(cp.id, 'DRAW');
                setTimeout(() => {
                    let hSum = cp.hand.reduce((s,c)=>s+c.val,0);
                    if (hSum < 15 && discardPile.length > 0) {
                        processPlayerAction(cp.id, 'KANG');
                    } else {
                        // Bot ทิ้งไพ่แต้มสูงสุด
                        let maxVal = -1; let maxIdx = 0;
                        cp.hand.forEach((c, i) => { if(c.val > maxVal) { maxVal = c.val; maxIdx = i; } });
                        processPlayerAction(cp.id, 'DISCARD', maxIdx);
                        setTimeout(() => processPlayerAction(cp.id, 'END_TURN'), 1500);
                    }
                }, 1500);
            }
        }, 2000);
    }
}

// Handler สิ้นสุดเกม
function handleRoundEndByCards(winner) {
    gameState.status = 'TRANSITION'; clearInterval(turnTimerInterval);
    let winAmt = 40; let winners = [winner.id]; let losers = [];
    gameState.players.forEach(p => {
        if (!p.isOut && p.id !== winner.id) { p.points -= winAmt; winner.points += winAmt; losers.push(p.id); }
    });
    
    let det = `[PID:${winner.id}] ไพ่หมดมือ ชนะรับคนละ ${winAmt} เหรียญ`;
    let isEnd = gameState.players.filter(x => x.points > 0).length <= 1;
    
    syncStateToAll();
    showResultModal("🏆 ไพ่หมดมือ!", det, isEnd, winners, losers);
    guestConnections.forEach(c => c.send({ type: 'SHOW_RESULT', title: "🏆 ไพ่หมดมือ!", detail: det, isEnd, winners, losers }));
}

function handleKangEnd(caller, isKang25) {
    let callerSum = caller.hand.reduce((s,c)=>s+c.val,0);
    let minScore = callerSum;
    let isCallerWin = true;
    
    gameState.players.forEach(p => {
        if(!p.isOut && p.id !== caller.id) {
            let score = p.hand.reduce((s,c)=>s+c.val,0);
            if(score <= minScore) { minScore = score; isCallerWin = false; }
        }
    });

    let winAmt = isKang25 ? 40 : 20;
    let det = ""; let winners = []; let losers = [];

    if (isCallerWin) {
        winners.push(caller.id);
        det += `[PID:${caller.id}] ชนะแคง! (แต้ม ${callerSum})<br>`;
        gameState.players.forEach(p => {
            if(!p.isOut && p.id !== caller.id) { p.points -= winAmt; caller.points += winAmt; losers.push(p.id); det += `[PID:${p.id}] เสีย ${winAmt} เหรียญ<br>`; }
        });
    } else {
        det += `[PID:${caller.id}] ถูกแหกแคง! (แต้ม ${callerSum})<br>`;
        gameState.players.forEach(p => {
            if(!p.isOut && p.id !== caller.id) {
                let score = p.hand.reduce((s,c)=>s+c.val,0);
                if(score === minScore) { p.points += (winAmt * 2); caller.points -= (winAmt * 2); winners.push(p.id); losers.push(caller.id); det += `[PID:${p.id}] แหกสำเร็จ! รับ ${winAmt * 2} เหรียญ<br>`; }
            }
        });
    }

    let isEnd = gameState.players.filter(x => x.points > 0).length <= 1;
    syncStateToAll();
    showResultModal(isCallerWin ? "🎉 แคงสำเร็จ!" : "💥 แหกแคง!", det, isEnd, winners, losers);
    guestConnections.forEach(c => c.send({ type: 'SHOW_RESULT', title: (isCallerWin ? "🎉 แคงสำเร็จ!" : "💥 แหกแคง!"), detail: det, isEnd, winners, losers }));
}