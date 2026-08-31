import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { getDatabase, ref, update, remove, onValue, get } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-database.js";

// --- Web Audio API System ---
const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
const soundBuffers = {};
const soundNames = ['1', 'select', 'start', 'bgm', 'jua', 'turn', 'uno', 'skip', 'reverse', 'draw2', 'draw4', 'll', 'ww', 'wl', 'win', 'abc'];

let turnAudioEnqueued = 0; // ติดตามจำนวนคิวของเสียง turn ที่กำลังจะเล่น/กำลังเล่นอยู่
let pendingABC = false; // ร้องขอการเล่น abc รอไว้

async function initAudio() {
    for (let name of soundNames) {
        try {
            const response = await fetch(`audio/${name}.mp3`);
            const arrayBuffer = await response.arrayBuffer();
            soundBuffers[name] = await audioCtx.decodeAudioData(arrayBuffer);
        } catch (e) { console.warn('Audio load fail:', name); }
    }
}
initAudio();

let bgmNode = null;
function playSound(name, onEndedCb = null) {
    if(audioCtx.state === 'suspended') audioCtx.resume();

    const handleTurnEnded = () => {
        if (name === 'turn') {
            turnAudioEnqueued = Math.max(0, turnAudioEnqueued - 1);
            // หากไม่มีคิวเสียง turn ค้างอยู่ และมีคำสั่งขอเล่น abc ล่วงหน้า
            if (turnAudioEnqueued === 0 && pendingABC) {
                pendingABC = false;
                // ยืนยันอีกครั้งว่ารอบปัจจุบันยังคงเป็นของตัวผู้เล่นเองจริงๆ
                const currentPlayer = players[game.turnIndex];
                if (currentPlayer && currentPlayer.id === myPeerId && !currentPlayer.isBot) {
                    playSound('abc');
                }
            }
        }
    };

    if(!soundBuffers[name]) {
        handleTurnEnded();
        if(onEndedCb) onEndedCb();
        return null;
    }
    const source = audioCtx.createBufferSource();
    source.buffer = soundBuffers[name];
    source.connect(audioCtx.destination);
    source.start(0);
    
    source.onended = () => {
        handleTurnEnded();
        if (onEndedCb) {
            onEndedCb();
        }
    };
    if (name === 'bgm') {
        source.loop = true;
        bgmNode = source;
    }
    return source;
}

function stopBGM() {
    if (bgmNode) {
        try { bgmNode.stop(); } catch(e){}
        bgmNode = null;
    }
}

function broadcastSound(soundName, playSelectFirst = false) {
    playSoundEvent(soundName, playSelectFirst);
    if(isHost) {
        connections.forEach(c => { if(c.open) c.send({ type: 'playSound', soundName, playSelectFirst }); });
    }
}

function playSoundEvent(soundName, playSelectFirst = false) {
    if (playSelectFirst) {
        playSound('select', () => playSound(soundName));
    } else {
        playSound(soundName);
    }
}
// -----------------------------

const firebaseConfig = {
    apiKey: "AIzaSyDvcdgsyT5sDdYTYKIqetzNL9Be-MFC0l4",
    authDomain: "xo-game-134ec.firebaseapp.com",
    databaseURL: "https://xo-game-134ec-default-rtdb.asia-southeast1.firebasedatabase.app",
    projectId: "xo-game-134ec",
    storageBucket: "xo-game-134ec.firebasestorage.app",
    messagingSenderId: "318375224157",
    appId: "1:318375224157:web:2eb3510bf38e153cb77eb4"
};
const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

let peer = null;
let myPeerId = null;
let isHost = false;
let connections = [];
let hostConnection = null;
let currentRoomId = null;

let isJoiningRoom = false;
let isCreatingRoom = false;
let lastSyncTimestamp = 0; 

const MAX_PLAYERS = 6;
const COLORS = [
    { id: "red", name: "สีแดง", hex: "#f56462" },
    { id: "blue", name: "สีน้ำเงิน", hex: "#00c3e5" },
    { id: "green", name: "สีเขียว", hex: "#2ecc71" },
    { id: "yellow", name: "สีเหลือง", hex: "#f7e018" },
    { id: "purple", name: "สีม่วง", hex: "#9b59b6" },
    { id: "orange", name: "สีส้ม", hex: "#e67e22" }
];

function getMascotEmoji(colorId) {
    switch (colorId) {
        case 'red': return '🔥';
        case 'blue': return '💧';
        case 'green': return '🍃';
        case 'yellow': return '⚡';
        case 'purple': return '🔮';
        case 'orange': return '☀️';
        default: return '😊';
    }
}

let players = [];
let myColorId = null;

let game = {
    status: 'waiting',
    deck: [],
    discardPile: [],
    turnIndex: 0,
    lastTurnIndex: -1, 
    direction: 1,
    currentColor: null,
    playerStates: {}, 
    currentTurnStartTime: 0,
    challengeData: null
};

let myBotDelay = null;

function getPronounName(playerObj) {
    if (!playerObj) return "";
    const colorDef = COLORS.find(c => c.id === playerObj.colorId);
    const colorName = colorDef ? colorDef.name : "ไม่ทราบสี";
    if (playerObj.id === myPeerId) return `คุณ${colorName}`;
    if (playerObj.isBot) return `บอท${colorName}`;
    return `เพื่อน${colorName}`;
}

let politeQueue = [];
let assertiveQueue = [];
let isAnnouncingPolite = false;
let isAnnouncingAssertive = false;

function processQueue(type) {
    const isAssertive = type === 'assertive';
    const queue = isAssertive ? assertiveQueue : politeQueue;
    if (queue.length === 0) {
        if (isAssertive) isAnnouncingAssertive = false;
        else isAnnouncingPolite = false;
        return;
    }
    if (isAssertive) isAnnouncingAssertive = true;
    else isAnnouncingPolite = true;
    const text = queue.shift();
    const elId = isAssertive ? 'aria-assertive' : 'aria-polite';
    const el = document.getElementById(elId);
    el.textContent = ''; 
    setTimeout(() => {
        el.textContent = text;
        const waitTime = Math.max(1000, text.length * 60); 
        setTimeout(() => {
            el.textContent = ''; 
            processQueue(type); 
        }, waitTime);
    }, 50);
}

function announce(text, assertive = false) {
    if (!text) return;
    if (assertive) {
        assertiveQueue.push(text);
        if (!isAnnouncingAssertive) processQueue('assertive');
    } else {
        politeQueue.push(text);
        if (!isAnnouncingPolite) processQueue('polite');
    }
}

function switchScreen(screenId, focusId) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    document.getElementById(screenId).classList.add('active');
    if (focusId) {
        setTimeout(() => {
            const el = document.getElementById(focusId);
            if (el) el.focus();
        }, 100);
    }
}

function disableBtn(btnId, disabled = true) {
    const btn = document.getElementById(btnId);
    if (btn) btn.disabled = disabled;
}

function initPeer() {
    peer = new Peer({ debug: 2 });
    peer.on('open', id => { myPeerId = id; });
    peer.on('connection', conn => {
        if (isHost) {
            conn.on('open', () => {
                connections.push(conn);
                setupHostConnection(conn);
                syncLobby();
            });
        }
    });
}
initPeer();

document.getElementById('btn-show-rules').onclick = () => switchScreen('screen-rules', 'title-rules');
document.getElementById('btn-close-rules').onclick = () => switchScreen('screen-main', 'btn-show-rules');
document.getElementById('btn-leave-lobby').onclick = leaveLobby;

const roomsRef = ref(db, 'uno_rooms');
onValue(roomsRef, (snapshot) => {
    if (document.getElementById('screen-main').classList.contains('active')) {
        renderRoomList(snapshot.val());
    }
});

function renderRoomList(rooms) {
    const list = document.getElementById('room-list');
    if (!rooms) {
        list.innerHTML = '<li id="empty-room-msg" style="text-align: center; color: var(--text-muted);">ไม่มีห้องที่เปิดอยู่</li>';
        return;
    }
    const activeRooms = Object.entries(rooms).filter(([id, r]) => Date.now() - r.lastActive < 120000 && r.status === 'waiting' && r.currentPlayers < MAX_PLAYERS);
    
    if (activeRooms.length === 0) {
        list.innerHTML = '<li id="empty-room-msg" style="text-align: center; color: var(--text-muted);">ไม่มีห้องที่เปิดอยู่</li>';
        return;
    }
    const emptyMsg = document.getElementById('empty-room-msg');
    if (emptyMsg) emptyMsg.remove();
    const activeRoomIds = new Set(activeRooms.map(([id]) => id));
    Array.from(list.children).forEach(li => {
        if (li.dataset.roomId && !activeRoomIds.has(li.dataset.roomId)) {
            li.remove();
        }
    });
    activeRooms.forEach(([id, room]) => {
        const hostCountText = `มีผู้เล่น ${room.currentPlayers} / ${MAX_PLAYERS} คน`;
        let li = list.querySelector(`li[data-room-id="${id}"]`);
        if (li) {
            const btn = li.querySelector('button');
            if (btn) {
                btn.innerHTML = `<span style="font-weight:bold; font-size:18px;">ห้อง: ${id}</span> <span>${hostCountText}</span>`;
                btn.setAttribute('aria-label', `เข้าร่วมห้อง ${id} ${hostCountText}`);
            }
        } else {
            li = document.createElement('li');
            li.dataset.roomId = id;
            li.style.margin = '10px 0';
            const btn = document.createElement('button');
            btn.style.width = '100%';
            btn.style.maxWidth = '100%';
            btn.style.textAlign = 'left';
            btn.style.display = 'flex';
            btn.style.justifyContent = 'space-between';
            btn.style.alignItems = 'center';
            btn.style.padding = '15px';
            btn.innerHTML = `<span style="font-weight:bold; font-size:18px;">ห้อง: ${id}</span> <span>${hostCountText}</span>`;
            btn.setAttribute('aria-label', `เข้าร่วมห้อง ${id} ${hostCountText}`);
            btn.onclick = () => {
                if (isJoiningRoom) return;
                isJoiningRoom = true;
                btn.disabled = true;
                btn.style.opacity = '0.5';
                joinRoom(id, room.hostPeerId);
            };
            li.appendChild(btn);
            list.appendChild(li);
        }
    });
}

document.getElementById('btn-create-room').onclick = async () => {
    if (isCreatingRoom) return;
    if (!myPeerId) {
        announce('ระบบกำลังเตรียมพร้อม กรุณารอสักครู่');
        return;
    }
    
    playSound('1');
    
    isCreatingRoom = true;
    disableBtn('btn-create-room', true);
    const btn = document.getElementById('btn-create-room');
    btn.style.opacity = '0.5';

    try {
        let nextIdNum = 1;
        const metaSnapshot = await get(ref(db, 'uno_metadata/last_room_id'));
        if (metaSnapshot.exists()) {
            nextIdNum = metaSnapshot.val() + 1;
            if (nextIdNum > 99999) nextIdNum = 1;
        }
        await update(ref(db, 'uno_metadata'), { last_room_id: nextIdNum });
        
        currentRoomId = "uno" + String(nextIdNum).padStart(5, '0');
        
        isHost = true;
        myColorId = getAvailableColors()[0];
        players = [{ id: myPeerId, name: "Host (คุณ)", isBot: false, colorId: myColorId, connection: null }];
        
        await update(ref(db, `uno_rooms/${currentRoomId}`), {
            hostPeerId: myPeerId, status: 'waiting', currentPlayers: 1, lastActive: Date.now()
        });

        setInterval(() => {
            if (isHost && currentRoomId && game.status === 'waiting') {
                update(ref(db, `uno_rooms/${currentRoomId}`), { lastActive: Date.now() });
            }
        }, 3000);

        enterLobby();
    } catch (err) {
        console.error(err);
        isCreatingRoom = false;
        disableBtn('btn-create-room', false);
        btn.style.opacity = '1';
        announce('สร้างห้องไม่สำเร็จ กรุณาลองใหม่');
    }
};

function joinRoom(roomId, hostPeerId) {
    if (!myPeerId) {
        announce('ระบบกำลังเตรียมพร้อม กรุณารอสักครู่');
        isJoiningRoom = false;
        return;
    }
    
    isHost = false; currentRoomId = roomId;
    announce('กำลังเชื่อมต่อไปยัง Host...');
    hostConnection = peer.connect(hostPeerId, { reliable: true });
    
    hostConnection.on('error', (err) => {
        console.error(err);
        announce('การเชื่อมต่อล้มเหลว', true);
        leaveLobby();
    });
    
    hostConnection.on('open', () => {
        hostConnection.send({ type: 'joinReq', peerId: myPeerId });
        enterLobby();
        hostConnection.on('data', handleClientData);
        hostConnection.on('close', handleHostDisconnect);
    });
}

function leaveLobby() {
    stopBGM();
    if (isHost) {
        if (currentRoomId) remove(ref(db, `uno_rooms/${currentRoomId}`));
        connections.forEach(c => c.close());
    } else if (hostConnection) {
        hostConnection.close();
    }
    currentRoomId = null; isHost = false; players = []; game.status = 'waiting';
    
    isJoiningRoom = false;
    isCreatingRoom = false;
    lastSyncTimestamp = 0;
    
    const btnCreate = document.getElementById('btn-create-room');
    if (btnCreate) {
        disableBtn('btn-create-room', false);
        btnCreate.style.opacity = '1';
    }
    
    switchScreen('screen-main', 'btn-create-room');
    announce('ออกจากห้องแล้ว');
}

function enterLobby() {
    switchScreen('screen-lobby', 'title-lobby');
    document.getElementById('lobby-room-id').textContent = currentRoomId;
    document.getElementById('host-controls').style.display = isHost ? 'block' : 'none';
    renderLobby();
}

function renderLobby() {
    const colorContainer = document.getElementById('color-selection');
    colorContainer.innerHTML = '';
    COLORS.forEach(c => {
        const isTaken = players.some(p => p.colorId === c.id && p.id !== myPeerId);
        const btn = document.createElement('button');
        btn.className = `lobby-color-btn card-${c.id} ${myColorId === c.id ? 'selected-color' : ''} ${isTaken ? 'disabled' : ''}`;
        btn.setAttribute('aria-label', c.name + (isTaken ? ' (ถูกเลือกแล้ว)' : ''));
        btn.innerHTML = `<span aria-hidden="true">${getMascotEmoji(c.id)}</span>`;
        if (myColorId === c.id) {
            btn.setAttribute('aria-pressed', 'true');
        }
        btn.onclick = () => {
            if (!isTaken) {
                myColorId = c.id;
                if (isHost) {
                    players.find(p => p.id === myPeerId).colorId = myColorId;
                    syncLobby();
                } else {
                    hostConnection.send({ type: 'changeColor', colorId: myColorId });
                }
            }
        };
        colorContainer.appendChild(btn);
    });

    const list = document.getElementById('lobby-player-list');
    list.innerHTML = '';
    document.getElementById('lobby-player-count').textContent = players.length;
    
    players.forEach(p => {
        const li = document.createElement('li');
        li.style.padding = '10px'; li.style.background = 'rgba(0,0,0,0.3)'; li.style.marginBottom = '8px'; li.style.borderRadius = '8px';
        li.style.display = 'flex'; li.style.alignItems = 'center'; li.style.gap = '10px';
        const cDef = COLORS.find(c => c.id === p.colorId);
        li.innerHTML = `<div style="width:24px; height:24px; border-radius:50%; background: ${cDef?cDef.hex:'#fff'}; display:flex; justify-content:center; align-items:center; font-size:12px; border:1px solid #fff;"><span aria-hidden="true">${p.isBot?'🤖':getMascotEmoji(p.colorId)}</span></div> ${getPronounName(p)}`;
        list.appendChild(li);
    });

    if (isHost) {
        const botCount = players.filter(p => p.isBot).length;
        document.getElementById('bot-count-display').textContent = botCount;
        disableBtn('btn-add-bot', players.length >= MAX_PLAYERS);
        disableBtn('btn-remove-bot', botCount === 0);
        disableBtn('btn-start-game', players.length < 2);
    }
}

function getAvailableColors() {
    const used = players.map(p => p.colorId);
    return COLORS.map(c => c.id).filter(id => !used.includes(id));
}

function setupHostConnection(conn) {
    conn.on('data', data => {
        if (data.type === 'joinReq') {
            conn.customPeerId = data.peerId; 
            if (players.length >= MAX_PLAYERS) return;
            players.push({ id: data.peerId, name: `Player`, isBot: false, colorId: getAvailableColors()[0], connection: conn });
            syncLobby();
            update(ref(db, `uno_rooms/${currentRoomId}`), { currentPlayers: players.length });
            broadcastGameEvent({ type: 'playerJoined', playerId: data.peerId });
            broadcastSound('select');
        }
        else if (data.type === 'changeColor') {
            const peerId = conn.customPeerId || conn.peer;
            const p = players.find(x => x.id === peerId);
            if (p && !players.some(x => x.colorId === data.colorId && x.id !== peerId)) {
                p.colorId = data.colorId; syncLobby();
            }
        }
        else if (data.type === 'action') {
            handlePlayerAction(conn.customPeerId || conn.peer, data.action, data.payload);
        }
    });
    conn.on('close', () => handleClientDisconnect(conn.customPeerId || conn.peer));
}

document.getElementById('btn-add-bot').onclick = () => {
    if (players.length < MAX_PLAYERS) {
        const colorId = getAvailableColors()[0];
        const botId = 'bot_' + Date.now();
        const colorDef = COLORS.find(c => c.id === colorId);
        players.push({ id: botId, name: `Bot`, isBot: true, colorId: colorId });
        announce(`เพิ่มบอท${colorDef.name} 1 ตัวแล้ว`, true);
        syncLobby();
        broadcastSound('select');
    }
};
document.getElementById('btn-remove-bot').onclick = () => {
    const botIdx = players.slice().reverse().findIndex(p => p.isBot);
    if (botIdx !== -1) {
        const actualIdx = players.length - 1 - botIdx;
        const removedBot = players[actualIdx];
        const colorDef = COLORS.find(c => c.id === removedBot.colorId);
        players.splice(actualIdx, 1);
        announce(`ลดบอท${colorDef.name} 1 ตัวแล้ว`, true);
        syncLobby();
        broadcastSound('select');
    }
};

function syncLobby() {
    if (!isHost) return;
    renderLobby();
    const safePlayers = players.map(p => ({ id: p.id, name: p.name, isBot: p.isBot, colorId: p.colorId }));
    connections.forEach(c => { if(c.open) c.send({ type: 'lobbySync', players: safePlayers }); });
}

function broadcastGameState() {
    if (!isHost) return;
    const safePlayers = players.map(p => ({ id: p.id, name: p.name, isBot: p.isBot, colorId: p.colorId }));
    const safeGame = {
        status: game.status,
        deckCount: game.deck.length,
        topDiscard: game.discardPile[game.discardPile.length - 1],
        discardPile: game.discardPile,
        turnIndex: game.turnIndex,
        direction: game.direction,
        currentColor: game.currentColor,
        challengeData: game.challengeData,
        players: safePlayers,
        playerStates: {},
        syncTimestamp: Date.now(),
        currentTurnStartTime: game.currentTurnStartTime
    };
    
    players.forEach(p => {
        safeGame.playerStates[p.id] = {
            cardCount: game.playerStates[p.id].hand.length,
            declaredUNO: game.playerStates[p.id].declaredUNO,
            score: game.playerStates[p.id].score,
            hasDrawn: game.playerStates[p.id].hasDrawn
        };
    });

    const activeConnections = connections.filter(c => c.open === true);
    activeConnections.forEach(c => {
        const targetPeerId = c.customPeerId || c.peer;
        const targetState = game.playerStates[targetPeerId];
        if (targetState) {
            const pState = { ...safeGame, myHand: [...targetState.hand] };
            c.send({ type: 'gameSync', game: pState });
        }
    });
    renderGame({ ...safeGame, myHand: [...game.playerStates[myPeerId].hand] });
    processTurnTimer();
}

function showUnoEffect() {
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    const overlay = document.getElementById('anim-overlay');
    const textEl = document.getElementById('anim-text');
    textEl.textContent = 'UNO!';
    textEl.className = 'anim-uno'; 
    overlay.style.display = 'flex';
    setTimeout(() => { 
        overlay.style.display = 'none'; 
        textEl.className = 'anim-text';
    }, 1500);
}

// --- Visual Animation Helpers (Presentation Layer Only) ---
function animateLocalDraw() {
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    try {
        const fxLayer = document.getElementById('visual-fx-layer');
        const deckEl = document.getElementById('deck-container');
        const handContainer = document.getElementById('my-cards-container');
        if (!fxLayer || !deckEl || !handContainer) return;

        const deckRect = deckEl.getBoundingClientRect();
        const handRect = handContainer.getBoundingClientRect();

        const clone = document.createElement('div');
        clone.className = 'uno-card deck-pile anim-card-clone';
        clone.style.position = 'fixed';
        clone.style.left = `${deckRect.left + deckRect.width/2 - 40}px`;
        clone.style.top = `${deckRect.top + deckRect.height/2 - 60}px`;
        clone.style.zIndex = '9995';
        clone.style.transition = 'all 0.6s cubic-bezier(0.2, 0.9, 0.3, 1)';
        clone.style.transform = 'scale(0.8) rotate(-10deg)';
        clone.style.pointerEvents = 'none';
        clone.setAttribute('aria-hidden', 'true');
        clone.innerHTML = '<span class="deck-pile-top" style="font-size:20px; color:#fff;">UNO</span>';

        fxLayer.appendChild(clone);

        const targetX = handRect.right - 80 > handRect.left ? handRect.right - 80 : handRect.left + handRect.width/2 - 40;
        const targetY = handRect.top + 10;

        requestAnimationFrame(() => {
            clone.style.left = `${targetX}px`;
            clone.style.top = `${targetY}px`;
            clone.style.transform = 'scale(1.05) rotate(5deg)';
        });

        setTimeout(() => {
            clone.style.opacity = '0';
        }, 550);

        setTimeout(() => {
            if (clone.parentNode) clone.parentNode.removeChild(clone);
        }, 650);
    } catch(e) {}
}

function animateRemoteDraw(playerId) {
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    try {
        const fxLayer = document.getElementById('visual-fx-layer');
        const deckEl = document.getElementById('deck-container');
        const charEl = document.querySelector(`.character-card[data-player-id="${playerId}"]`);
        if (!fxLayer || !deckEl || !charEl) return;

        const deckRect = deckEl.getBoundingClientRect();
        const charRect = charEl.getBoundingClientRect();

        const clone = document.createElement('div');
        clone.className = 'uno-card deck-pile anim-card-clone';
        clone.style.position = 'fixed';
        clone.style.left = `${deckRect.left + deckRect.width/2 - 25}px`;
        clone.style.top = `${deckRect.top + deckRect.height/2 - 35}px`;
        clone.style.width = '50px';
        clone.style.height = '70px';
        clone.style.zIndex = '9995';
        clone.style.transition = 'all 0.6s cubic-bezier(0.4, 0, 0.2, 1)';
        clone.style.transform = 'scale(0.7)';
        clone.style.pointerEvents = 'none';
        clone.setAttribute('aria-hidden', 'true');
        clone.innerHTML = '<span style="font-size:12px; color:var(--focus-ring); font-weight:bold;">UNO</span>';

        fxLayer.appendChild(clone);

        const targetX = charRect.left + charRect.width/2 - 25;
        const targetY = charRect.top + charRect.height/2 - 35;

        requestAnimationFrame(() => {
            clone.style.left = `${targetX}px`;
            clone.style.top = `${targetY}px`;
            clone.style.transform = 'scale(0.4) rotate(15deg)';
            clone.style.opacity = '0.9';
        });

        setTimeout(() => {
            clone.style.transform = 'scale(0.1)';
            clone.style.opacity = '0';
            charEl.classList.add('char-bounce-reaction');
            setTimeout(() => charEl.classList.remove('char-bounce-reaction'), 400);
        }, 550);

        setTimeout(() => {
            if (clone.parentNode) clone.parentNode.removeChild(clone);
        }, 650);
    } catch(e) {}
}

function animateLocalPlay(card, selectedColor) {
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    try {
        const fxLayer = document.getElementById('visual-fx-layer');
        const handContainer = document.getElementById('my-cards-container');
        const discardEl = document.getElementById('discard-container');
        if (!fxLayer || !handContainer || !discardEl) return;

        const handRect = handContainer.getBoundingClientRect();
        const discardRect = discardEl.getBoundingClientRect();

        const clone = document.createElement('div');
        const displayColor = (card.color === 'wild' && selectedColor) ? selectedColor : card.color;
        clone.className = `uno-card card-${displayColor} anim-card-clone`;
        clone.style.position = 'fixed';
        clone.style.left = `${handRect.left + handRect.width/2 - 40}px`;
        clone.style.top = `${handRect.top - 20}px`;
        clone.style.zIndex = '9995';
        clone.style.transition = 'all 0.5s cubic-bezier(0.2, 0.8, 0.2, 1)';
        clone.style.transform = 'scale(1.15) rotate(0deg)';
        clone.style.pointerEvents = 'none';
        clone.setAttribute('aria-hidden', 'true');

        let inner = card.value;
        if (card.type === 'skip') inner = '<span class="card-special-text">Skip</span>';
        else if (card.type === 'reverse') inner = '<span class="card-special-text">Reverse</span>';
        else if (card.type === 'draw2') inner = '<span class="card-special-text">Draw Two</span>';
        else if (card.type === 'wild') inner = '<span class="card-wild-text">Wild</span>';
        else if (card.type === 'wild4') inner = '<span class="card-wild-text">Wild Draw Four</span>';
        clone.innerHTML = inner;

        fxLayer.appendChild(clone);

        const targetX = discardRect.left + discardRect.width/2 - 40;
        const targetY = discardRect.top + discardRect.height/2 - 60;

        requestAnimationFrame(() => {
            clone.style.left = `${targetX}px`;
            clone.style.top = `${targetY}px`;
            clone.style.transform = 'scale(1) rotate(4deg)';
        });

        setTimeout(() => {
            clone.style.opacity = '0';
            triggerDiscardImpact(card.type, displayColor);
        }, 450);

        setTimeout(() => {
            if (clone.parentNode) clone.parentNode.removeChild(clone);
        }, 550);
    } catch(e) {}
}

function animateRemotePlay(playerId, card, selectedColor) {
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    try {
        const fxLayer = document.getElementById('visual-fx-layer');
        const charEl = document.querySelector(`.character-card[data-player-id="${playerId}"]`);
        const discardEl = document.getElementById('discard-container');
        if (!fxLayer || !charEl || !discardEl) return;

        const charRect = charEl.getBoundingClientRect();
        const discardRect = discardEl.getBoundingClientRect();

        const clone = document.createElement('div');
        const displayColor = (card.color === 'wild' && selectedColor) ? selectedColor : card.color;
        clone.className = `uno-card card-${displayColor} anim-card-clone`;
        clone.style.position = 'fixed';
        clone.style.left = `${charRect.left + charRect.width/2 - 30}px`;
        clone.style.top = `${charRect.top + charRect.height/2 - 45}px`;
        clone.style.width = '60px';
        clone.style.height = '90px';
        clone.style.zIndex = '9995';
        clone.style.transition = 'all 0.55s cubic-bezier(0.2, 0.8, 0.2, 1)';
        clone.style.transform = 'scale(0.6) rotate(-10deg)';
        clone.style.pointerEvents = 'none';
        clone.setAttribute('aria-hidden', 'true');

        let inner = card.value;
        if (card.type === 'skip') inner = '<span class="card-special-text">Skip</span>';
        else if (card.type === 'reverse') inner = '<span class="card-special-text">Reverse</span>';
        else if (card.type === 'draw2') inner = '<span class="card-special-text">Draw Two</span>';
        else if (card.type === 'wild') inner = '<span class="card-wild-text">Wild</span>';
        else if (card.type === 'wild4') inner = '<span class="card-wild-text">Wild Draw Four</span>';
        clone.innerHTML = inner;

        fxLayer.appendChild(clone);

        const targetX = discardRect.left + discardRect.width/2 - 40;
        const targetY = discardRect.top + discardRect.height/2 - 60;

        requestAnimationFrame(() => {
            clone.style.left = `${targetX}px`;
            clone.style.top = `${targetY}px`;
            clone.style.width = '80px';
            clone.style.height = '120px';
            clone.style.transform = 'scale(1) rotate(-3deg)';
        });

        setTimeout(() => {
            clone.style.opacity = '0';
            triggerDiscardImpact(card.type, displayColor);
        }, 500);

        setTimeout(() => {
            if (clone.parentNode) clone.parentNode.removeChild(clone);
        }, 600);
    } catch(e) {}
}

function triggerDiscardImpact(type, color) {
    try {
        const fxLayer = document.getElementById('visual-fx-layer');
        const discardEl = document.getElementById('discard-container');
        if (!fxLayer || !discardEl) return;

        const rect = discardEl.getBoundingClientRect();
        const centerX = rect.left + rect.width / 2;
        const centerY = rect.top + rect.height / 2;

        if (type === 'skip') {
            const ring = document.createElement('div');
            ring.className = 'fx-skip-ring';
            ring.style.left = `${centerX - 50}px`;
            ring.style.top = `${centerY - 50}px`;
            ring.innerHTML = '⛔';
            ring.setAttribute('aria-hidden', 'true');
            fxLayer.appendChild(ring);
            setTimeout(() => ring.remove(), 700);
        } else if (type === 'reverse') {
            const sweep = document.createElement('div');
            sweep.className = 'fx-reverse-sweep';
            sweep.style.left = `${centerX - 60}px`;
            sweep.style.top = `${centerY - 60}px`;
            sweep.innerHTML = '🔄';
            sweep.setAttribute('aria-hidden', 'true');
            fxLayer.appendChild(sweep);
            setTimeout(() => sweep.remove(), 800);
        } else if (type === 'draw2') {
            const burst = document.createElement('div');
            burst.className = 'fx-draw-burst';
            burst.style.left = `${centerX - 40}px`;
            burst.style.top = `${centerY - 40}px`;
            burst.innerHTML = '+2';
            burst.setAttribute('aria-hidden', 'true');
            fxLayer.appendChild(burst);
            setTimeout(() => burst.remove(), 800);
        } else if (type === 'wild4' || type === 'wild') {
            for (let i = 0; i < 12; i++) {
                const p = document.createElement('div');
                p.className = 'fx-particle';
                p.style.left = `${centerX}px`;
                p.style.top = `${centerY}px`;
                p.style.backgroundColor = ['#f56462', '#00c3e5', '#2ecc71', '#f7e018'][i % 4];
                p.setAttribute('aria-hidden', 'true');
                const angle = (i / 12) * Math.PI * 2;
                const dist = 60 + Math.random() * 40;
                p.style.setProperty('--tx', `${Math.cos(angle) * dist}px`);
                p.style.setProperty('--ty', `${Math.sin(angle) * dist}px`);
                fxLayer.appendChild(p);
                setTimeout(() => p.remove(), 750);
            }
        }
    } catch(e) {}
}

function triggerUnoCinematicAnimation(playerId) {
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    try {
        const fxLayer = document.getElementById('visual-fx-layer');
        if (!fxLayer) return;

        const player = players.find(p => p.id === playerId);
        const colorDef = player ? COLORS.find(c => c.id === player.colorId) : null;
        const colorHex = colorDef ? colorDef.hex : '#f56462';
        const mascot = player ? (player.isBot ? '🤖' : getMascotEmoji(player.colorId)) : '😊';

        const overlay = document.createElement('div');
        overlay.className = 'fx-uno-cinematic-overlay';
        overlay.setAttribute('aria-hidden', 'true');

        overlay.innerHTML = `
            <div class="fx-uno-cinematic-content">
                <div class="fx-uno-mascot" style="background-color: ${colorHex};">
                    <span>${mascot}</span>
                </div>
                <div class="fx-uno-title">UNO!</div>
            </div>
        `;

        fxLayer.appendChild(overlay);

        setTimeout(() => {
            overlay.classList.add('fade-out');
            setTimeout(() => overlay.remove(), 200);
        }, 850);
    } catch(e) {}
}

function handleClientData(data) {
    if (data.type === 'lobbySync') {
        players = data.players;
        const me = players.find(p => p.id === myPeerId);
        if (me) myColorId = me.colorId;
        renderLobby();
    } else if (data.type === 'gameSync') {
        if (data.game.syncTimestamp) {
            if (data.game.syncTimestamp < lastSyncTimestamp) return;
            lastSyncTimestamp = data.game.syncTimestamp;
        }
        if (data.game.players) players = data.game.players;
        if ((game.status === 'waiting' || game.status === 'ended') && data.game.status === 'playing') {
            switchScreen('screen-game', 'top-status-bar');
            announce('เริ่มรอบใหม่แล้ว!', true);
        }
        game = { ...game, ...data.game };
        renderGame(data.game);
    } else if (data.type === 'gameEvent') {
        handleGameEvent(data.eventData);
    } else if (data.type === 'announce') {
        announce(data.message, true);
        showAnimOverlay(data.message);
    } else if (data.type === 'unoEffect') {
        showUnoEffect();
    } else if (data.type === 'endGame') {
        game.status = 'ended'; 
        doWinAnimation(() => {
            showResult(data.winnerId, data.scores, data.matchOver);
        });
    } else if (data.type === 'playSound') {
        playSoundEvent(data.soundName, data.playSelectFirst);
    } else if (data.type === 'startAnim') {
        doStartAnimation(() => {});
    }
}

function handleHostDisconnect() { announce('Host หลุดการเชื่อมต่อ', true); leaveLobby(); }
function handleClientDisconnect(peerId) {
    if (!isHost) return;
    const p = players.find(x => x.id === peerId);
    if (p && !p.isBot) {
        const disconnectMsg = `${getPronounName(p)}หลุดการเชื่อมต่อ เปลี่ยนเป็นบอทแล้ว`;
        p.isBot = true;
        broadcastAnnounce(disconnectMsg, true);
        syncLobby();
        if (game.status === 'playing') broadcastGameState();
    }
}

function doStartAnimation(callback) {
    stopBGM();
    playSound('start');
    const animDiv = document.createElement('div');
    animDiv.id = 'start-anim-screen';
    animDiv.setAttribute('aria-hidden', 'true');
    animDiv.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:radial-gradient(circle at center, #1b382b 0%, #000000 90%);z-index:10000;display:flex;justify-content:center;align-items:center;color:#f7e018;font-size:5rem;font-weight:900;text-shadow: 0 0 30px #f7e018, 0 0 60px #f56462;';
    document.body.appendChild(animDiv);
    
    const steps = [
        { time: 200, text: 'Uno' },
        { time: 1200, text: '3' },
        { time: 2200, text: '2' },
        { time: 3200, text: '1' },
        { time: 4200, text: 'Enjoy' }
    ];
    
    steps.forEach(step => {
        setTimeout(() => {
            animDiv.textContent = step.text;
            announce(step.text, true); 
        }, step.time);
    });
    
    setTimeout(() => {
        animDiv.remove();
        if(callback) callback();
    }, 5000);
}

function doWinAnimation(callback) {
    stopBGM();
    const animDiv = document.createElement('div');
    animDiv.id = 'win-anim-screen';
    animDiv.setAttribute('aria-hidden', 'true');
    animDiv.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:radial-gradient(circle at center, #111827 0%, #000000 100%);z-index:10000;display:flex;justify-content:center;align-items:center;color:#ffffff;font-size:3rem;font-weight:bold;text-align:center;text-shadow: 0 0 20px #f7e018;';
    document.body.appendChild(animDiv);
    
    setTimeout(() => { playSound('win'); }, 500);

    const steps = [
        { time: 200, text: 'Uno' },
        { time: 1200, text: 'กำลังทำการสรุปผล' },
        { time: 2500, text: 'ได้แก่' }
    ];
    
    steps.forEach(step => {
        setTimeout(() => {
            animDiv.textContent = step.text;
            announce(step.text, true);
        }, step.time);
    });
    
    setTimeout(() => {
        animDiv.remove();
        if(callback) callback();
    }, 3500);
}

document.getElementById('btn-start-game').onclick = () => {
    disableBtn('btn-start-game');
    document.getElementById('btn-start-game').style.display = 'none'; 
    if (isHost && players.length >= 2) {
        remove(ref(db, `uno_rooms/${currentRoomId}`));
        connections.forEach(c => { if(c.open) c.send({ type: 'startAnim' }); });
        doStartAnimation(() => {
            initUNOGame();
            setTimeout(() => {
                playSound('bgm');
                connections.forEach(c => { if(c.open) c.send({ type: 'playSound', soundName: 'bgm' }); });
            }, 200);
        });
    }
};

function generateDeck() {
    const deck = [];
    const colors = ['red', 'blue', 'green', 'yellow'];
    colors.forEach(c => {
        deck.push({ color: c, type: 'number', value: 0 });
        for(let i=1; i<=9; i++) {
            deck.push({ color: c, type: 'number', value: i });
            deck.push({ color: c, type: 'number', value: i });
        }
        for(let i=0; i<2; i++) {
            deck.push({ color: c, type: 'skip', value: 'Skip' });
            deck.push({ color: c, type: 'reverse', value: 'Reverse' });
            deck.push({ color: c, type: 'draw2', value: 'Draw Two' });
        }
    });
    for(let i=0; i<4; i++) {
        deck.push({ color: 'wild', type: 'wild', value: 'Wild' });
        deck.push({ color: 'wild', type: 'wild4', value: 'Wild Draw Four' });
    }
    return deck.sort(() => Math.random() - 0.5);
}

function initUNOGame() {
    game.status = 'playing';
    game.deck = generateDeck();
    game.discardPile = [];
    game.direction = 1;
    game.turnIndex = 0;
    game.lastTurnIndex = -1;
    game.challengeData = null;
    
    players.forEach(p => {
        const existingScore = game.playerStates[p.id] ? game.playerStates[p.id].score : 0;
        game.playerStates[p.id] = { hand: [], declaredUNO: false, score: existingScore, hasDrawn: false };
        for(let i=0; i<7; i++) drawCardHost(p.id);
    });

    let topCard = game.deck.pop();
    while(topCard.type === 'wild4') { game.deck.unshift(topCard); topCard = game.deck.pop(); }
    game.discardPile.push(topCard);
    game.currentColor = topCard.color === 'wild' ? ['red','blue','green','yellow'][Math.floor(Math.random()*4)] : topCard.color;

    if (topCard.type === 'skip') game.turnIndex = (game.turnIndex + 1) % players.length;
    else if (topCard.type === 'reverse') { game.direction = -1; game.turnIndex = (players.length - 1) % players.length; }
    else if (topCard.type === 'draw2') { drawCardHost(players[game.turnIndex].id, 2); game.turnIndex = (game.turnIndex + 1) % players.length; }

    switchScreen('screen-game', 'top-status-bar');
    broadcastAnnounce('เกมเริ่มแล้ว!');
    game.currentTurnStartTime = Date.now();
    broadcastGameState();
}

function drawCardHost(playerId, count = 1) {
    for(let i=0; i<count; i++) {
        if (game.deck.length === 0) {
            if (game.discardPile.length <= 1) break;
            const top = game.discardPile.pop();
            game.deck = game.discardPile.sort(() => Math.random() - 0.5);
            game.discardPile = [top];
            broadcastAnnounce('สับกองทิ้งใหม่');
        }
        game.playerStates[playerId].hand.push(game.deck.pop());
        game.playerStates[playerId].declaredUNO = false;
    }
}

function nextTurn(steps = 1) {
    game.turnIndex = (game.turnIndex + (steps * game.direction) + players.length * 10) % players.length;
    game.currentTurnStartTime = Date.now();
    game.challengeData = null;
    
    players.forEach(p => {
        const s = game.playerStates[p.id];
        s.hasDrawn = false;
        if (s.hand.length === 1 && !s.declaredUNO && p.id !== players[game.turnIndex].id) {
            broadcastGameEvent({ type: 'forgotUno', playerId: p.id });
            drawCardHost(p.id, 2);
        }
    });
    broadcastGameState();
}

function isValidPlay(card, hand) {
    if (card.color === 'wild') return true;
    const top = game.discardPile[game.discardPile.length - 1];
    if (top.color === 'wild') {
        return card.color === game.currentColor;
    }
    return card.color === game.currentColor || card.value === top.value;
}

function handleGameEvent(data) {
    let msg = '';
    let assertive = false;
    switch (data.type) {
        case 'forgotUno': {
            const p = players.find(x => x.id === data.playerId);
            if (p) { msg = `${getPronounName(p)} ลืมประกาศ UNO โดนปรับจั่ว 2 ใบ`; assertive = true; }
            break;
        }
        case 'declareUno': {
            const p = players.find(x => x.id === data.playerId);
            if (p) { msg = `${getPronounName(p)} ประกาศ UNO!`; assertive = true; playSound('uno'); }
            triggerUnoCinematicAnimation(data.playerId);
            break;
        }
        case 'challengeSuccess': {
            const p = players.find(x => x.id === data.playerId);
            if (p) msg = `Challenge สำเร็จ! ${getPronounName(p)} โดนปรับจั่ว 4 ใบ`;
            playSound('ww');
            break;
        }
        case 'challengeFailed': {
            const p = players.find(x => x.id === data.playerId);
            if (p) msg = `Challenge พลาด! ${getPronounName(p)} โดนปรับจั่ว 6 ใบ`;
            playSound('wl');
            break;
        }
        case 'challengeDeclined': {
            const p = players.find(x => x.id === data.playerId);
            if (p) msg = `${getPronounName(p)} ไม่ยอมรับ Challenge โดนจั่ว 4 ใบ`;
            playSound('ll');
            break;
        }
        case 'drawCard': {
            const p = players.find(x => x.id === data.playerId);
            if (p) {
                msg = `${getPronounName(p)} จั่วการ์ด 1 ใบ`;
                if (p.isBot || data.isAuto) playSound('jua');
            }
            if (data.playerId === myPeerId) {
                animateLocalDraw();
            } else {
                animateRemoteDraw(data.playerId);
            }
            break;
        }
        case 'passTurn': {
            const p = players.find(x => x.id === data.playerId);
            if (p) msg = `${getPronounName(p)} จบตา`;
            turnAudioEnqueued++;
            playSound('turn');
            break;
        }
        case 'winGame': {
            const p = players.find(x => x.id === data.winnerId);
            if (p) { msg = `จบเกม! ผู้ชนะคือ ${getPronounName(p)}`; assertive = true; }
            break;
        }
        case 'playCard': {
            const p = players.find(x => x.id === data.playerId);
            if (!p) break;
            msg = `${getPronounName(p)} วางการ์ด ${getCardARIA(data.card)}`;
            
            if (data.playerId === myPeerId) {
                animateLocalPlay(data.card, data.selectedColor);
            } else {
                animateRemotePlay(data.playerId, data.card, data.selectedColor);
            }

            if (data.card.color === 'wild' && data.selectedColor) {
                const colorNames = {red:'แดง', blue:'น้ำเงิน', green:'เขียว', yellow:'เหลือง'};
                msg += ` เปลี่ยนเป็นสี${colorNames[data.selectedColor]}`;
            }

            const sequence = [];
            if (p.isBot || data.isAuto) sequence.push('select');

            if (data.effect === 'skip') {
                sequence.push('skip');
                const target = players.find(x => x.id === data.targetId);
                if (target) msg += `, ส่งผลให้ ${getPronounName(target)} ถูกใช้การ์ด skip ข้ามตา`;
            } else if (data.effect === 'reverse') {
                sequence.push('reverse');
                msg += `, ทำการกลับทิศทางการเล่น`;
            } else if (data.effect === 'draw2') {
                sequence.push('draw2');
                const target = players.find(x => x.id === data.targetId);
                if (target) msg += `, ส่งผลให้ ${getPronounName(target)} โดนบังคับจั่ว 2 ใบและถูกข้ามตา`;
            } else if (data.effect === 'wild4') {
                sequence.push('draw4');
                const target = players.find(x => x.id === data.targetId);
                if (target) msg += `, ส่งผลให้ ${getPronounName(target)} ถูกใช้ Wild Draw Four ต้องเลือกว่าจะ Challenge หรือไม่`;
            }
            
            if (!data.isWin) {
                sequence.push('turn');
                turnAudioEnqueued++;
            }

            const playSeq = (arr) => {
                if (!arr || arr.length === 0) return;
                const s = arr.shift();
                playSound(s, () => playSeq(arr));
            };
            playSeq(sequence);

            break;
        }
        case 'playerJoined': {
            const p = players.find(x => x.id === data.playerId);
            if (p) msg = `${getPronounName(p)} เข้าร่วมห้องสำเร็จ`;
            break;
        }
    }

    if (msg) {
        announce(msg, assertive);
        if (!isHost) showAnimOverlay(msg); 
    }
}

function broadcastGameEvent(eventData) {
    handleGameEvent(eventData);
    if(isHost) connections.forEach(c => { if (c.open) c.send({ type: 'gameEvent', eventData: eventData }); });
}

function handlePlayerAction(peerId, action, payload = null) {
    if (!isHost || game.status !== 'playing') return;
    const currentPlayer = players[game.turnIndex];
    
    if (game.challengeData && game.challengeData.victimId === peerId) {
        if (action === 'challenge_yes') {
            const attackerId = game.challengeData.attackerId;
            const hasMatchingColor = game.playerStates[attackerId].hand.some(c => c.color === game.challengeData.prevColor);
            if (hasMatchingColor) {
                broadcastGameEvent({ type: 'challengeSuccess', playerId: attackerId });
                drawCardHost(attackerId, 4);
            } else {
                broadcastGameEvent({ type: 'challengeFailed', playerId: peerId });
                drawCardHost(peerId, 6);
            }
            nextTurn();
        } else if (action === 'challenge_no') {
            broadcastGameEvent({ type: 'challengeDeclined', playerId: peerId });
            drawCardHost(peerId, 4); 
            nextTurn();
        }
        return;
    }

    if (currentPlayer.id !== peerId) {
        if (action === 'announce_uno' && game.playerStates[peerId].hand.length === 2 && !game.playerStates[peerId].declaredUNO) {
            game.playerStates[peerId].declaredUNO = true;
            broadcastGameEvent({ type: 'declareUno', playerId: peerId });
            
            if (isHost) connections.forEach(c => { if (c.open) c.send({ type: 'unoEffect' }); });
            showUnoEffect();
            
            broadcastGameState();
        }
        return;
    }

    const state = game.playerStates[peerId];

    if (action === 'announce_uno' && state.hand.length === 2 && !state.declaredUNO) {
        state.declaredUNO = true;
        broadcastGameEvent({ type: 'declareUno', playerId: peerId });
        
        if (isHost) connections.forEach(c => { if (c.open) c.send({ type: 'unoEffect' }); });
        showUnoEffect();
        
        broadcastGameState();
        return;
    }

    if (action === 'play') {
        const card = state.hand[payload.cardIndex];
        if (isValidPlay(card, state.hand)) {
            state.hand.splice(payload.cardIndex, 1);
            game.discardPile.push(card);
            game.currentColor = card.color !== 'wild' ? card.color : payload.selectedColor;

            const isWin = state.hand.length === 0;
            const isAuto = payload && payload.isAuto;

            if (state.hand.length === 0) { 
                broadcastGameEvent({
                    type: 'playCard',
                    playerId: peerId,
                    card: card,
                    selectedColor: payload.selectedColor,
                    effect: 'none',
                    isWin: isWin,
                    isAuto: isAuto
                });
                handleWin(peerId); 
                return; 
            }

            if (card.type === 'skip') {
                const nextPlayer = players[(game.turnIndex + game.direction + players.length) % players.length];
                broadcastGameEvent({
                    type: 'playCard',
                    playerId: peerId,
                    card: card,
                    selectedColor: payload.selectedColor,
                    effect: 'skip',
                    targetId: nextPlayer.id,
                    isWin: isWin,
                    isAuto: isAuto
                });
                nextTurn(2);
            }
            else if (card.type === 'reverse') {
                broadcastGameEvent({
                    type: 'playCard',
                    playerId: peerId,
                    card: card,
                    selectedColor: payload.selectedColor,
                    effect: 'reverse',
                    isWin: isWin,
                    isAuto: isAuto
                });
                game.direction = game.direction === 1 ? -1 : 1;
                nextTurn(players.length === 2 ? 2 : 1);
            }
            else if (card.type === 'draw2') { 
                const nextPlayer = players[(game.turnIndex + game.direction + players.length) % players.length];
                broadcastGameEvent({
                    type: 'playCard',
                    playerId: peerId,
                    card: card,
                    selectedColor: payload.selectedColor,
                    effect: 'draw2',
                    targetId: nextPlayer.id,
                    isWin: isWin,
                    isAuto: isAuto
                });
                drawCardHost(players[(game.turnIndex + game.direction + players.length) % players.length].id, 2); 
                nextTurn(2); 
            }
            else if (card.type === 'wild4') {
                const nextPlayer = players[(game.turnIndex + game.direction + players.length) % players.length];
                broadcastGameEvent({
                    type: 'playCard',
                    playerId: peerId,
                    card: card,
                    selectedColor: payload.selectedColor,
                    effect: 'wild4',
                    targetId: nextPlayer.id,
                    isWin: isWin,
                    isAuto: isAuto
                });
                game.challengeData = { victimId: nextPlayer.id, attackerId: peerId, prevColor: game.currentColor };
                broadcastGameState();
            }
            else {
                broadcastGameEvent({
                    type: 'playCard',
                    playerId: peerId,
                    card: card,
                    selectedColor: payload.selectedColor,
                    effect: 'none',
                    isWin: isWin,
                    isAuto: isAuto
                });
                nextTurn(1);
            }
        }
    } else if (action === 'draw') {
        if (!state.hasDrawn) {
            broadcastGameEvent({ type: 'drawCard', playerId: peerId, isAuto: payload && payload.isAuto });
            drawCardHost(peerId, 1);
            state.hasDrawn = true;
            
            const validCardsCount = state.hand.filter(c => isValidPlay(c, state.hand)).length;
            if (validCardsCount === 0) {
                broadcastGameEvent({ type: 'passTurn', playerId: peerId, isAuto: payload && payload.isAuto });
                nextTurn(1);
            } else {
                broadcastGameState();
            }
        }
    } else if (action === 'pass') {
        if (state.hasDrawn) {
            broadcastGameEvent({ type: 'passTurn', playerId: currentPlayer.id, isAuto: payload && payload.isAuto });
            nextTurn(1);
        }
    }
}

function handleWin(winnerId) {
    broadcastGameEvent({ type: 'winGame', winnerId: winnerId });
    let score = 0;
    players.forEach(p => {
        if(p.id !== winnerId) {
            game.playerStates[p.id].hand.forEach(c => {
                if (c.type === 'number') score += c.value;
                else if (c.type === 'wild' || c.type === 'wild4') score += 50;
                else score += 20;
            });
        }
    });
    
    game.playerStates[winnerId].score += score;
    if (game.playerStates[winnerId].score >= 200) {
        game.matchOver = true;
    } else {
        game.matchOver = false;
    }
    
    if (isHost) {
        broadcastGameState(); 
        connections.forEach(c => { if (c.open) c.send({ type: 'endGame', winnerId, scores: score, matchOver: game.matchOver }); });
        game.status = 'ended'; 
        doWinAnimation(() => {
            showResult(winnerId, score, game.matchOver);
        });
    }
}

function broadcastAnnounce(msg, assertive = false) {
    announce(msg, assertive);
    if(isHost) connections.forEach(c => { if (c.open) c.send({ type: 'announce', message: msg }); });
}

function processTurnTimer() {
    if (!isHost || game.status !== 'playing') return;
    const currentPlayer = players[game.turnIndex];
    
    if (game.challengeData) {
        if (players.find(p => p.id === game.challengeData.victimId).isBot) {
            setTimeout(() => handlePlayerAction(game.challengeData.victimId, 'challenge_no'), 2500);
        }
        return;
    }

    if (currentPlayer.isBot) {
        clearTimeout(myBotDelay);
        const delay = 3000 + Math.random() * 2000;
        myBotDelay = setTimeout(() => {
            if (game.status !== 'playing' || players[game.turnIndex].id !== currentPlayer.id) return;
            
            const state = game.playerStates[currentPlayer.id];
            const validCards = state.hand.map((c, i) => ({card: c, index: i})).filter(x => isValidPlay(x.card, state.hand));
            
            if (state.hand.length === 2 && validCards.length > 0 && !state.declaredUNO) {
                handlePlayerAction(currentPlayer.id, 'announce_uno');
            }

            if (validCards.length > 0) {
                const playChoice = validCards[Math.floor(Math.random() * validCards.length)];
                handlePlayerAction(currentPlayer.id, 'play', { cardIndex: playChoice.index, selectedColor: ['red', 'blue', 'green', 'yellow'][Math.floor(Math.random() * 4)], isAuto: true });
            } else {
                if (!state.hasDrawn) {
                    handlePlayerAction(currentPlayer.id, 'draw', { isAuto: true });
                } else {
                    handlePlayerAction(currentPlayer.id, 'pass', { isAuto: true });
                }
            }
        }, delay);
    }
}

window.turnTimerInterval = setInterval(() => {
    if(!isHost || game.status !== 'playing') return;
    const currentPlayer = players[game.turnIndex];
    if(!currentPlayer || currentPlayer.isBot) return;
    
    if(game.challengeData) {
        if(Date.now() - game.currentTurnStartTime > 40000) {
             handlePlayerAction(game.challengeData.victimId, 'challenge_no');
        }
        return;
    }

    if(Date.now() - game.currentTurnStartTime > 40000) {
        const state = game.playerStates[currentPlayer.id];
        const validCards = state.hand.map((c, i) => ({card: c, index: i})).filter(x => isValidPlay(x.card, state.hand));
        
        if (state.hand.length === 2 && validCards.length > 0 && !state.declaredUNO) {
            handlePlayerAction(currentPlayer.id, 'announce_uno');
        }
        if (validCards.length > 0) {
            const playChoice = validCards[Math.floor(Math.random() * validCards.length)];
            handlePlayerAction(currentPlayer.id, 'play', { cardIndex: playChoice.index, selectedColor: ['red', 'blue', 'green', 'yellow'][Math.floor(Math.random() * 4)], isAuto: true });
        } else {
            if (!state.hasDrawn) {
                handlePlayerAction(currentPlayer.id, 'draw', { isAuto: true });
            } else {
                handlePlayerAction(currentPlayer.id, 'pass', { isAuto: true });
            }
        }
    }
}, 1000);

function getCardARIA(card) {
    if (!card) return '';
    const cName = { red: 'แดง ', blue: 'น้ำเงิน ', green: 'เขียว ', yellow: 'เหลือง ', wild: '' }[card.color];
    let vName = card.value;
    if (card.type === 'skip') vName = 'Skip';
    else if (card.type === 'reverse') vName = 'Reverse';
    else if (card.type === 'draw2') vName = 'Draw Two';
    else if (card.type === 'wild') vName = 'Wild';
    else if (card.type === 'wild4') vName = 'Wild Draw Four';
    return `${cName}${vName}`.trim();
}

function renderCardHTML(card, index = -1, isPlayable = false) {
    let inner = '';
    if (card.type === 'number') inner = card.value;
    else if (card.type === 'skip') inner = '<span class="card-special-text">Skip</span>';
    else if (card.type === 'reverse') inner = '<span class="card-special-text">Reverse</span>';
    else if (card.type === 'draw2') inner = '<span class="card-special-text">Draw Two</span>';
    else if (card.type === 'wild') inner = '<span class="card-wild-text">Wild</span>';
    else if (card.type === 'wild4') inner = '<span class="card-wild-text">Wild Draw Four</span>';

    const btn = document.createElement(index >= 0 ? 'button' : 'div');
    btn.className = `uno-card card-${card.color} ${!isPlayable && index >= 0 ? 'disabled' : ''}`;
    btn.innerHTML = inner;
    
    if (index >= 0) {
        btn.setAttribute('aria-label', `${getCardARIA(card)} ${isPlayable ? 'ลงได้' : 'ลงไม่ได้'}`);
        if (isPlayable) btn.onclick = () => onCardClicked(index, card);
        else btn.setAttribute('aria-disabled', 'true');
    }
    return btn;
}

function toggleMainGameUI(show) {
    const topBar = document.getElementById('top-status-bar');
    const midArea = document.querySelector('.middle-area');
    const bottomArea = document.getElementById('my-cards-container').parentElement;
    
    if (show) {
        topBar.style.display = 'flex';
        midArea.style.display = 'flex';
        bottomArea.style.display = 'flex';
        topBar.removeAttribute('aria-hidden');
        midArea.removeAttribute('aria-hidden');
        bottomArea.removeAttribute('aria-hidden');
    } else {
        topBar.style.display = 'none';
        midArea.style.display = 'none';
        bottomArea.style.display = 'none';
        topBar.setAttribute('aria-hidden', 'true');
        midArea.setAttribute('aria-hidden', 'true');
        bottomArea.setAttribute('aria-hidden', 'true');
    }
}

function renderGame(gameState) {
    const turnPlayerObj = players[gameState.turnIndex];
    if (turnPlayerObj) {
        const headingEl = document.getElementById('current-turn-heading');
        headingEl.textContent = `ถึงรอบ${getPronounName(turnPlayerObj)}`;
        
        // Color heading visual based on turn player's color
        const cDef = COLORS.find(c => c.id === turnPlayerObj.colorId);
        if (cDef) {
            headingEl.style.backgroundColor = cDef.hex;
            if (turnPlayerObj.colorId === 'yellow') {
                headingEl.style.color = '#0f172a';
                headingEl.style.textShadow = '0 1px 2px rgba(255,255,255,0.5)';
                headingEl.style.boxShadow = '0 0 15px rgba(247, 224, 24, 0.6)';
            } else {
                headingEl.style.color = '#ffffff';
                headingEl.style.textShadow = '0 1px 2px rgba(0,0,0,0.8)';
                headingEl.style.boxShadow = `0 0 15px ${cDef.hex}`;
            }
        }
    }

    // Update top bar
    const topBar = document.getElementById('top-status-bar');
    if(topBar) topBar.innerHTML = '';
    
    const dirIndicator = document.createElement('div');
    dirIndicator.className = 'direction-indicator';
    dirIndicator.textContent = gameState.direction === 1 ? '➡️' : '⬅️';
    if(topBar) topBar.appendChild(dirIndicator);

    players.forEach((p, idx) => {
        const charCard = document.createElement('div');
        charCard.className = `character-card ${idx === gameState.turnIndex ? 'is-turn' : ''}`;
        charCard.dataset.playerId = p.id;
        
        const charState = gameState.playerStates[p.id];
        const cardCount = charState ? charState.cardCount : 0;
        
        charCard.innerHTML = `
            <div class="char-cards-count">${cardCount} ใบ</div>
            <div class="char-avatar" style="background-color: ${COLORS.find(c => c.id === p.colorId)?.hex || '#fff'}">
                ${p.isBot ? '🤖' : getMascotEmoji(p.colorId)}
            </div>
            <div class="char-score">${p.name}</div>
        `;
        if(topBar) topBar.appendChild(charCard);
    });

    // Update Deck & Discard
    const discardPile = document.getElementById('discard-pile');
    if(discardPile) discardPile.innerHTML = '';
    if (gameState.topDiscard && discardPile) {
        discardPile.appendChild(renderCardHTML(gameState.topDiscard, -1, false));
        const ariaDiscard = document.getElementById('discard-aria-label');
        if(ariaDiscard) ariaDiscard.textContent = `การ์ดบนสุดคือ ${getCardARIA(gameState.topDiscard)}`;
    }
    
    const deckCountEl = document.getElementById('deck-count-visual');
    if(deckCountEl) deckCountEl.textContent = gameState.deckCount;
    const deckAria = document.getElementById('deck-aria-label');
    if(deckAria) deckAria.textContent = `การ์ดในกองเหลือ ${gameState.deckCount} ใบ`;

    // Hand
    const myState = gameState.playerStates[myPeerId];
    const myHandContainer = document.getElementById('my-cards-container');
    if(myHandContainer) {
        myHandContainer.innerHTML = '';
        if (gameState.myHand.length > 8) {
            myHandContainer.classList.add('compact');
        } else {
            myHandContainer.classList.remove('compact');
        }

        const isMyTurn = players[gameState.turnIndex].id === myPeerId && !gameState.challengeData;
        
        gameState.myHand.forEach((card, idx) => {
            const isPlayable = isMyTurn && isValidPlay(card, gameState.myHand);
            myHandContainer.appendChild(renderCardHTML(card, idx, isPlayable));
        });
    }

    // Challenge Area
    const challengeArea = document.getElementById('challenge-area');
    if(challengeArea) {
        if (gameState.challengeData && gameState.challengeData.victimId === myPeerId) {
            toggleMainGameUI(false);
            challengeArea.style.display = 'block';
            document.getElementById('btn-challenge-yes').onclick = () => handlePlayerAction(myPeerId, 'challenge_yes');
            document.getElementById('btn-challenge-no').onclick = () => handlePlayerAction(myPeerId, 'challenge_no');
        } else {
            toggleMainGameUI(true);
            challengeArea.style.display = 'none';
        }
    }

    // Action Buttons
    const btnDraw = document.getElementById('btn-draw');
    const btnPass = document.getElementById('btn-pass');
    const btnUno = document.getElementById('btn-uno');
    const isMyTurn = players[gameState.turnIndex].id === myPeerId && !gameState.challengeData;

    if(btnDraw && btnPass && btnUno && myState) {
        btnDraw.disabled = !isMyTurn || myState.hasDrawn;
        btnPass.disabled = !isMyTurn || !myState.hasDrawn;
        
        const hasPlayableCard = gameState.myHand.some(c => isValidPlay(c, gameState.myHand));
        
        if (isMyTurn && !myState.hasDrawn && !hasPlayableCard) {
            btnDraw.classList.add('is-turn'); // Highlight draw
        } else {
            btnDraw.classList.remove('is-turn');
        }

        if (myState.hand && myState.hand.length === 2 && !myState.declaredUNO) {
            btnUno.disabled = false;
            btnUno.onclick = () => handlePlayerAction(myPeerId, 'announce_uno');
        } else {
            btnUno.disabled = true;
        }

        btnDraw.onclick = () => handlePlayerAction(myPeerId, 'draw');
        btnPass.onclick = () => handlePlayerAction(myPeerId, 'pass');
    }
    
    if (gameState.status === 'playing') {
        if (gameState.lastTurnIndex !== gameState.turnIndex && players[gameState.turnIndex].id === myPeerId && !players[gameState.turnIndex].isBot) {
            if (turnAudioEnqueued === 0) {
                playSound('abc');
            } else {
                pendingABC = true;
            }
        }
        game.lastTurnIndex = gameState.turnIndex;
    }
}

let pendingWildCardIndex = -1;

window.onCardClicked = function(index, card) {
    if (card.color === 'wild') {
        pendingWildCardIndex = index;
        document.getElementById('color-picker-modal').style.display = 'flex';
        document.getElementById('color-picker-title').focus();
    } else {
        handlePlayerAction(myPeerId, 'play', { cardIndex: index, selectedColor: null });
    }
};

window.selectWildColor = function(color) {
    document.getElementById('color-picker-modal').style.display = 'none';
    if (pendingWildCardIndex >= 0) {
        handlePlayerAction(myPeerId, 'play', { cardIndex: pendingWildCardIndex, selectedColor: color });
        pendingWildCardIndex = -1;
    }
};

function showAnimOverlay(msg) {
    const overlay = document.getElementById('anim-overlay');
    if(!overlay) return;
    const textEl = document.getElementById('anim-text');
    textEl.textContent = msg;
    overlay.style.display = 'flex';
    setTimeout(() => { overlay.style.display = 'none'; }, 2000);
}

function showResult(winnerId, score, matchOver) {
    switchScreen('screen-result', 'title-result');
    const winner = players.find(p => p.id === winnerId);
    const winText = document.getElementById('winner-text');
    if(winText) winText.innerHTML = `<span aria-hidden="true">🏆</span> ผู้ชนะ: ${getPronounName(winner)}`;
    
    const hostControls = document.getElementById('result-host-controls');
    if(hostControls) {
        if (isHost) {
            hostControls.style.display = 'flex';
            const btnPlayAgain = document.getElementById('btn-play-again');
            if(btnPlayAgain) {
                btnPlayAgain.onclick = () => {
                    if (matchOver) {
                        players.forEach(p => {
                            if (game.playerStates[p.id]) game.playerStates[p.id].score = 0;
                        });
                    }
                    initUNOGame();
                };
            }
        } else {
            hostControls.style.display = 'none';
        }
    }
    
    const statusBar = document.getElementById('result-status-bar');
    if(statusBar) {
        statusBar.innerHTML = '';
        players.forEach(p => {
            const s = game.playerStates[p.id];
            const scoreText = s ? s.score : 0;
            statusBar.innerHTML += `<div style="padding: 5px 10px; background: rgba(255,255,255,0.1); border-radius: 8px;">${getPronounName(p)}: ${scoreText} คะแนน</div>`;
        });
    }
}

const btnRefreshAudio = document.getElementById('btn-refresh-audio');
if(btnRefreshAudio) {
    btnRefreshAudio.onclick = () => {
        if(audioCtx.state === 'suspended') audioCtx.resume();
        playSound('select');
    };
}
