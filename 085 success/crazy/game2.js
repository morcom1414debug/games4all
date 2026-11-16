import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { getDatabase, ref, update, remove, onValue, get } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-database.js";

// Firebase Config from provided file
const firebaseConfig = {
	apiKey: "AIzaSyDvcdgsyT5sDdYTYKIqetzNL9Be-MFC0l4",
	authDomain: "xo-game-134ec.firebaseapp.com",
	databaseURL: "https://xo-game-134ec-default-rtdb.asia-southeast1.firebasedatabase.app",
	projectId: "xo-game-134ec",
	storageBucket: "xo-game-134ec.firebasestorage.app",
	messagingSenderId: "318375224157",
	appId: "1:318375224157:web:9d953686dea05222b77eb4"
};
const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

// --- Audio System (Unchanged Timing & Logic) ---
const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
const soundBuffers = {};
const soundNames = ['1', 'select', 'start', 'bgm', 'jua', 'turn', 'uno', 'win', 'hit'];
const audioQueue = [];
let isAudioPlaying = false;
let bgmNode = null;

async function initAudio() {
	for (let name of soundNames) {
		try {
			const response = await fetch(`audio/${name}.mp3`);
			if(response.ok) {
				const arrayBuffer = await response.arrayBuffer();
				soundBuffers[name] = await audioCtx.decodeAudioData(arrayBuffer);
			}
		} catch (e) { console.warn('Audio load fail:', name); }
	}
}
initAudio();

function processAudioQueue() {
	if (isAudioPlaying || audioQueue.length === 0) return;
	isAudioPlaying = true;
	const task = audioQueue.shift();
	const handleEnd = () => { if (task.onEndedCb) task.onEndedCb(); isAudioPlaying = false; processAudioQueue(); };
	
	if (!soundBuffers[task.name]) { handleEnd(); return; }
	const source = audioCtx.createBufferSource();
	source.buffer = soundBuffers[task.name];
	source.connect(audioCtx.destination);
	source.onended = handleEnd;
	try { source.start(0); } catch (e) { handleEnd(); }
}

function playSound(name, onEndedCb = null) {
	if(audioCtx.state === 'suspended') audioCtx.resume();
	if (name === 'bgm') {
		if(!soundBuffers[name]) return null;
		const source = audioCtx.createBufferSource();
		source.buffer = soundBuffers[name];
		source.connect(audioCtx.destination);
		source.loop = true;
		source.start(0);
		bgmNode = source;
		return source;
	}
	audioQueue.push({ name, onEndedCb });
	processAudioQueue();
	return null;
}

function stopBGM() { if (bgmNode) { try { bgmNode.stop(); } catch(e){} bgmNode = null; } }

function broadcastSound(soundName) {
	playSound(soundName);
	if(isHost) connections.forEach(c => { if(c.open) c.send({ type: 'playSound', soundName }); });
}

// --- ARIA System (Unchanged Focus & Timing) ---
let politeQueue = [];
let assertiveQueue = [];
let isAnnouncingPolite = false;
let isAnnouncingAssertive = false;

function processQueue(type) {
	const isAssertive = type === 'assertive';
	const queue = isAssertive ? assertiveQueue : politeQueue;
	if (queue.length === 0) { isAssertive ? isAnnouncingAssertive = false : isAnnouncingPolite = false; return; }
	
	isAssertive ? isAnnouncingAssertive = true : isAnnouncingPolite = true;
	const text = queue.shift();
	const el = document.getElementById(isAssertive ? 'aria-assertive' : 'aria-polite');
	el.textContent = ''; 
	setTimeout(() => {
		el.textContent = text;
		setTimeout(() => { el.textContent = ''; processQueue(type); }, Math.max(500, text.length * 30));
	}, 30);
}

function announce(text, assertive = false) {
	if (!text) return;
	if (assertive) { assertiveQueue.push(text); if (!isAnnouncingAssertive) processQueue('assertive'); }
	else { politeQueue.push(text); if (!isAnnouncingPolite) processQueue('polite'); }
}

function broadcastAnnounce(msg, assertive = false) {
	announce(msg, assertive);
	logEvent(msg);
	if(isHost) connections.forEach(c => { if (c.open) c.send({ type: 'announce', message: msg, assertive }); });
}

function logEvent(msg) {
	const logEl = document.getElementById('recent-event-log');
	if (logEl) logEl.textContent = msg;
}

// --- Game State & Multiplayer ---
let peer = null;
let myPeerId = null;
let myName = "";
let isHost = false;
let connections = [];
let hostConnection = null;
let currentRoomId = null;
let isJoiningRoom = false;
let isCreatingRoom = false;
let heartbeatInterval = null;

const MAX_PLAYERS = 4;
// Bot Names Reference from Domino
const BOT_NAMES = [
	'สายฟ้า', 'เจ้าป่า', 'ดาวเหนือ', 'ขุนพล', 'จอมทัพ', 
	'พายุ', 'ฟีนิกซ์', 'นักรบ', 'เสือดำ', 'ราชัน',
	'มังกร', 'ภูผา', 'ทะเล', 'วายุ', 'หมอก', 
	'ตะวัน', 'จันทรา', 'แสงดาว', 'เพชร', 'โชคดี'
];

let players = [];
let game = {
	status: 'waiting',
	deck: [],
    discardPile: [],
	turnIndex: 0,
	direction: 1,
	playerStates: {}, 
	matchOver: false,
    activeSuit: null
};

function initPeer() {
	peer = new Peer({ debug: 2 });
	peer.on('open', id => { myPeerId = id; });
	peer.on('connection', conn => {
		if (isHost) {
			conn.on('open', () => { connections.push(conn); setupHostConnection(conn); syncLobby(); });
		}
	});
}
initPeer();

function switchScreen(screenId, focusId) {
	document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
	document.getElementById(screenId).classList.add('active');
	if (focusId) setTimeout(() => { const el = document.getElementById(focusId); if (el) el.focus(); }, 100);
}

const nameInput = document.getElementById('player-name-input');
const btnConfirmName = document.getElementById('btn-confirm-name');

nameInput.addEventListener('input', () => {
	if (nameInput.value.trim().length > 0) { btnConfirmName.disabled = false; }
	else { btnConfirmName.disabled = true; }
});

nameInput.addEventListener('keydown', (e) => {
	if ((e.key === 'Enter' || e.key === 'Return') && !btnConfirmName.disabled) {
		btnConfirmName.onclick();
	}
});

btnConfirmName.onclick = () => {
	myName = nameInput.value.trim();
	document.getElementById('display-player-name').textContent = myName;
	switchScreen('screen-main', 'title-main');
	playSound('select');
};

document.getElementById('btn-show-rules').onclick = () => { playSound('select'); switchScreen('screen-rules', 'title-rules'); };
document.getElementById('btn-close-rules').onclick = () => { playSound('select'); switchScreen('screen-main', 'btn-show-rules'); };

// --- Firebase Rooms ---
const roomsRef = ref(db, 'crazy_rooms');
onValue(roomsRef, (snapshot) => {
	if (document.getElementById('screen-main').classList.contains('active')) renderRoomList(snapshot.val());
});

function renderRoomList(rooms) {
	const list = document.getElementById('room-list');
	const activeRooms = rooms ? Object.entries(rooms).filter(([id, r]) => Date.now() - r.lastActive < 120000 && r.status === 'waiting' && r.currentPlayers < MAX_PLAYERS) : [];
	if (activeRooms.length === 0) {
		list.innerHTML = '<li id="empty-room-msg" style="text-align: center; color: var(--text-muted);">ไม่มีห้องที่เปิดอยู่</li>';
		return;
	}
	const emptyMsg = document.getElementById('empty-room-msg');
	if (emptyMsg) emptyMsg.remove();
	const existingItems = Array.from(list.querySelectorAll('li[data-room-id]'));
	const activeRoomIds = activeRooms.map(([id]) => id);

	existingItems.forEach(li => {
		const roomId = li.getAttribute('data-room-id');
		if (!activeRoomIds.includes(roomId)) { li.remove(); }
	});

	activeRooms.forEach(([id, room]) => {
		let li = list.querySelector(`li[data-room-id="${id}"]`);
		if (li) {
			const btn = li.querySelector('button');
			if (btn) {
				btn.innerHTML = `<span>ห้อง: ${id}</span> <span>${room.currentPlayers}/${MAX_PLAYERS} คน</span>`;
				btn.setAttribute('aria-label', `เข้าร่วมห้อง ${id} มีผู้เล่น ${room.currentPlayers} จาก ${MAX_PLAYERS} คน`);
			}
		} else {
			li = document.createElement('li'); 
			li.style.margin = '10px 0';
			li.setAttribute('data-room-id', id);
			const btn = document.createElement('button');
			btn.style.width = '100%'; btn.style.maxWidth = '100%'; btn.style.display = 'flex'; btn.style.justifyContent = 'space-between';
			btn.innerHTML = `<span>ห้อง: ${id}</span> <span>${room.currentPlayers}/${MAX_PLAYERS} คน</span>`;
			btn.setAttribute('aria-label', `เข้าร่วมห้อง ${id} มีผู้เล่น ${room.currentPlayers} จาก ${MAX_PLAYERS} คน`);
			btn.onclick = () => {
				if (isJoiningRoom) return;
				isJoiningRoom = true; btn.disabled = true; playSound('1');
				joinRoom(id, room.hostPeerId);
			};
			li.appendChild(btn); 
			list.appendChild(li);
		}
	});
}

document.getElementById('btn-create-room').onclick = async () => {
	if (isCreatingRoom) return;
	if (!myPeerId) { announce('ระบบกำลังเตรียมพร้อม กรุณารอสักครู่'); return; }
	playSound('1');
	isCreatingRoom = true; document.getElementById('btn-create-room').disabled = true;

	try {
		let nextIdNum = 1;
		const metaSnapshot = await get(ref(db, 'crazy_metadata/last_room_id'));
		if (metaSnapshot.exists()) {
			nextIdNum = metaSnapshot.val() + 1;
			if (nextIdNum > 99999) nextIdNum = 1;
		}
		await update(ref(db, 'crazy_metadata'), { last_room_id: nextIdNum });
		
        // Room Prefix: Crazy
		currentRoomId = "Crazy" + String(nextIdNum).padStart(5, '0');
		isHost = true;
		players = [{ id: myPeerId, name: myName, isBot: false }];
		
		clearInterval(heartbeatInterval);
		heartbeatInterval = setInterval(hostHeartbeatCheck, 3000);

		await update(ref(db, `crazy_rooms/${currentRoomId}`), { hostPeerId: myPeerId, status: 'waiting', currentPlayers: 1, lastActive: Date.now() });

		setInterval(() => { if (isHost && currentRoomId && game.status === 'waiting') update(ref(db, `crazy_rooms/${currentRoomId}`), { lastActive: Date.now() }); }, 3000);
		enterLobby();
		document.getElementById('lobby-ready-msg').style.display = 'block';
		announce('ห้องพร้อมแล้ว รอเพื่อนหรือเพิ่มบอทได้ทันที');
	} catch (err) {
		isCreatingRoom = false; document.getElementById('btn-create-room').disabled = false; announce('สร้างห้องไม่สำเร็จ');
	}
};

function joinRoom(roomId, hostPeerId) {
	if (!myPeerId) { announce('ระบบกำลังเตรียมพร้อม กรุณารอสักครู่'); isJoiningRoom = false; return; }
	isHost = false; currentRoomId = roomId;
	announce('กำลังเชื่อมต่อไปยัง Host...');
	hostConnection = peer.connect(hostPeerId, { reliable: true });
	hostConnection.on('error', () => { announce('การเชื่อมต่อล้มเหลว', true); leaveLobby(); });
	hostConnection.on('open', () => {
		hostConnection.send({ type: 'joinReq', peerId: myPeerId, name: myName });
		enterLobby();
		hostConnection.on('data', handleClientData);
		hostConnection.on('close', () => { announce('Host หลุดการเชื่อมต่อ', true); leaveLobby(); });
	});
}

window.leaveLobby = function() {
	stopBGM();
	clearInterval(heartbeatInterval);
	if (isHost && currentRoomId) { remove(ref(db, `crazy_rooms/${currentRoomId}`)); connections.forEach(c => c.close()); }
	else if (hostConnection) { hostConnection.close(); }
	currentRoomId = null; isHost = false; players = []; game.status = 'waiting';
	isJoiningRoom = false; isCreatingRoom = false;
	document.getElementById('btn-create-room').disabled = false;
	document.getElementById('lobby-ready-msg').style.display = 'none';
	switchScreen('screen-main', 'title-main');
	announce('ออกจากห้องแล้ว');
}
document.getElementById('btn-leave-lobby').onclick = leaveLobby;

function enterLobby() {
	switchScreen('screen-lobby', 'title-lobby');
	document.getElementById('lobby-room-id').textContent = currentRoomId;
	document.getElementById('host-controls').style.display = isHost ? 'block' : 'none';
    document.getElementById('lobby-ready-msg').style.display = 'block';
	renderLobby();
}

function renderLobby() {
	const list = document.getElementById('lobby-player-list');
	list.innerHTML = '';
	document.getElementById('lobby-player-count').textContent = players.length;
	players.forEach(p => {
		const li = document.createElement('li');
		li.style.padding = '12px 15px'; li.style.background = 'rgba(255,255,255,0.1)'; li.style.marginBottom = '8px'; li.style.borderRadius = '12px';
		li.style.display = 'flex'; li.style.alignItems = 'center'; li.style.gap = '10px'; li.style.border = '1px solid rgba(255,255,255,0.05)';
		li.innerHTML = `<span aria-hidden="true" style="font-size: 20px;">${p.isBot?'🤖':'👤'}</span> <strong>${p.name}</strong> ${p.id === myPeerId ? '(คุณ)' : ''}`;
		list.appendChild(li);
	});
	if (isHost) {
		const botCount = players.filter(p => p.isBot).length;
		document.getElementById('bot-count-display').textContent = botCount;
		document.getElementById('btn-add-bot').disabled = (players.length >= MAX_PLAYERS);
		document.getElementById('btn-remove-bot').disabled = (botCount === 0);
		document.getElementById('btn-start-game').disabled = (players.length < 2);
	}
}

function handleClientDisconnect(peerId) {
	const p = players.find(x => x.id === peerId);
	if (p && !p.isBot) {
		p.isBot = true; p.name = `บอท${p.name}`;
		broadcastAnnounce(`เพื่อน${p.name} หลุดการเชื่อมต่อ เปลี่ยนเป็นบอทแล้ว`, true);
		syncLobby(); 
		if (game.status === 'playing') {
			broadcastGameState();
			if (players[game.turnIndex].id === peerId) { processTurnLogic(); }
		}
	}
}

function hostHeartbeatCheck() {
	if (!isHost) return;
	const now = Date.now();
	connections.forEach(conn => {
		if (!conn.lastPing) conn.lastPing = now;
		if (conn.open) { conn.send({ type: 'ping' }); }
		if (now - conn.lastPing > 15000) {
			const peerId = conn.customPeerId || conn.peer;
			handleClientDisconnect(peerId);
			if (conn.open) conn.close();
			conn.lastPing = now; 
		}
	});
}

function setupHostConnection(conn) {
	conn.lastPing = Date.now();
	conn.on('data', data => {
		if (data.type === 'pong') {
			conn.lastPing = Date.now();
		} else if (data.type === 'joinReq') {
			conn.customPeerId = data.peerId; 
			if (players.length >= MAX_PLAYERS) return;
			let newName = data.name;
			let count = 1;
			while(players.some(p => p.name === newName)) { newName = `${data.name}(${count++})`; }
			players.push({ id: data.peerId, name: newName, isBot: false, connection: conn });
			syncLobby();
			update(ref(db, `crazy_rooms/${currentRoomId}`), { currentPlayers: players.length });
			broadcastSound('select');
			broadcastAnnounce(`${newName} เข้าร่วมห้องสำเร็จ`);
		} else if (data.type === 'action') {
			handlePlayerAction(conn.customPeerId || conn.peer, data.action, data.payload);
		}
	});
	conn.on('close', () => {
		const peerId = conn.customPeerId || conn.peer;
		handleClientDisconnect(peerId);
	});
}

document.getElementById('btn-add-bot').onclick = () => {
	if (players.length < MAX_PLAYERS) {
		const availableNames = BOT_NAMES.filter(n => !players.some(p => p.name === n));
		const randomIdx = Math.floor(Math.random() * availableNames.length);
		const botName = availableNames[randomIdx] || `บอท_${Date.now().toString().slice(-4)}`;
		players.push({ id: 'bot_' + Date.now(), name: botName, isBot: true });
		syncLobby(); broadcastSound('select');
		broadcastAnnounce(`เพิ่มบอท ${botName} เข้าห้องแล้ว`, true);
		if (currentRoomId) update(ref(db, `crazy_rooms/${currentRoomId}`), { currentPlayers: players.length });
	}
};
document.getElementById('btn-remove-bot').onclick = () => {
	const botIdx = players.slice().reverse().findIndex(p => p.isBot);
	if (botIdx !== -1) {
		const botToRemove = players[players.length - 1 - botIdx];
		players.splice(players.length - 1 - botIdx, 1);
		syncLobby(); broadcastSound('select');
		broadcastAnnounce(`ลด ${botToRemove.name} ออกจากห้องแล้ว`, true);
		if (currentRoomId) update(ref(db, `crazy_rooms/${currentRoomId}`), { currentPlayers: players.length });
	}
};

function syncLobby() {
	if (!isHost) return;
	renderLobby();
	const safePlayers = players.map(p => ({ id: p.id, name: p.name, isBot: p.isBot }));
	connections.forEach(c => { if(c.open) c.send({ type: 'lobbySync', players: safePlayers }); });
}

// --- CRAZY EIGHTS LOGIC ---
function generateDeck() {
	const deck = [];
    const suits = ['♥', '♦', '♣', '♠'];
    const ranks = ['A', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K'];
    suits.forEach(suit => {
        ranks.forEach(rank => {
            let colorName = (suit === '♥' || suit === '♦') ? 'แดง' : 'ดำ';
            let suitName = '';
            if(suit === '♥') suitName = 'โพแดง';
            else if(suit === '♦') suitName = 'ข้าวหลามตัด';
            else if(suit === '♣') suitName = 'ดอกจิก';
            else if(suit === '♠') suitName = 'โพดำ';
            let rankName = rank;
            if (rank === 'A') rankName = 'เอซ';
            else if (rank === 'J') rankName = 'แจ็ค';
            else if (rank === 'Q') rankName = 'แหม่ม';
            else if (rank === 'K') rankName = 'คิง';
            deck.push({ id: `${suit}${rank}`, suit, rank, name: `${rankName} ${suitName}` });
        });
    });
	return deck.sort(() => Math.random() - 0.5);
}

function renderCardHTML(card, index = -1, playable = false, isTop = false) {
	const btn = document.createElement((index >= 0 && !isTop) ? 'button' : 'div');
    const isRed = (card.suit === '♥' || card.suit === '♦');
	btn.className = `crazy-card ${isRed ? 'red-suit' : 'black-suit'} ${(index >= 0 && !playable && !isTop) ? 'disabled' : ''}`;
	btn.innerHTML = `<div class="card-top">${card.rank}</div><div class="card-center">${card.suit}</div><div class="card-bottom">${card.rank}</div>`;
	
	if (index >= 0 && !isTop) {
		btn.setAttribute('aria-label', `${card.name} ${playable ? 'ลงได้' : 'ลงไม่ได้'}`);
		if (playable) {
			btn.onclick = (e) => {
                if (card.rank === '8') {
                    // เปิด Modal เลือกดอก
                    pendingPlayIndex = index;
                    document.getElementById('suit-picker-modal').style.display = 'flex';
                } else {
				    sendAction('play', { index });
                }
			};
		} else {
			btn.setAttribute('aria-disabled', 'true');
		}
	} else if (isTop) {
        btn.setAttribute('aria-label', `ไพ่กองทิ้งคือ ${card.name}`);
    }
	return btn;
}

let pendingPlayIndex = -1;

document.querySelectorAll('.suit-btn').forEach(btn => {
    btn.onclick = (e) => {
        const selectedSuit = e.target.getAttribute('data-suit');
        document.getElementById('suit-picker-modal').style.display = 'none';
        sendAction('play', { index: pendingPlayIndex, activeSuit: selectedSuit });
        pendingPlayIndex = -1;
    };
});

document.getElementById('btn-start-game').onclick = () => {
	document.getElementById('btn-start-game').disabled = true;
	if (isHost && players.length >= 2) {
		remove(ref(db, `crazy_rooms/${currentRoomId}`));
		connections.forEach(c => { if(c.open) c.send({ type: 'startAnim' }); });
		doStartAnimation(() => {
			initGame();
			setTimeout(() => broadcastSound('bgm'), 500);
		});
	}
};

function doStartAnimation(callback) {
	stopBGM(); playSound('start');
	const animDiv = document.createElement('div');
	animDiv.setAttribute('aria-hidden', 'true');
	animDiv.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:radial-gradient(circle at center, #1b382b 0%, #000000 90%);z-index:10000;display:flex;justify-content:center;align-items:center;color:#fff;font-size:5rem;font-weight:900;text-shadow:0 0 30px #4CAF50,0 0 60px #2e7d32; flex-direction:column;';
	document.body.appendChild(animDiv);
	
	const steps = [
		{ time: 200, text: 'CRAZY EIGHTS' }, { time: 1200, text: '3' }, { time: 2200, text: '2' }, { time: 3200, text: '1' }, { time: 4200, text: 'Enjoy' }
	];
	steps.forEach(step => {
		setTimeout(() => { animDiv.textContent = step.text; announce(step.text, true); }, step.time);
	});
	setTimeout(() => { animDiv.remove(); if(callback) callback(); }, 5000);
}

function broadcastTurnStart() {
	if (!isHost) return;
	const currentPlayer = players[game.turnIndex];
    let topCard = game.discardPile[game.discardPile.length - 1];
    let msg = `ถึงรอบของ ${currentPlayer.name === myName ? 'คุณ' + myName : currentPlayer.name}`;
	broadcastAnnounce(msg);
}

function initGame() {
	game.status = 'playing'; game.matchOver = false;
	game.deck = generateDeck();
    game.discardPile = [];
	game.turnIndex = 0;
    game.activeSuit = null;
	
	players.forEach((p, i) => {
		game.playerStates[p.id] = { hand: [] };
		for(let k=0; k<5; k++) game.playerStates[p.id].hand.push(game.deck.pop());
	});

    let top = game.deck.pop();
    while(top.rank === '8') {
        game.deck.unshift(top);
        top = game.deck.pop();
    }
    game.discardPile.push(top);
    game.activeSuit = top.suit;

	switchScreen('screen-game', 'top-status-bar');
	broadcastGameState();
	
	setTimeout(() => {
		broadcastTurnStart();
		broadcastSound('turn');
		processTurnLogic();
	}, 1000);
}

function broadcastGameState() {
	if (!isHost) return;
	const safePlayers = players.map(p => ({ id: p.id, name: p.name, isBot: p.isBot }));
	const safeGame = {
		status: game.status,
		deckCount: game.deck.length,
        topCard: game.discardPile[game.discardPile.length - 1],
        activeSuit: game.activeSuit,
		turnIndex: game.turnIndex,
		players: safePlayers,
		playerStates: {}
	};
	players.forEach(p => {
		safeGame.playerStates[p.id] = { cardCount: game.playerStates[p.id].hand.length };
	});

	connections.forEach(c => {
		if(c.open) {
			const tId = c.customPeerId || c.peer;
			c.send({ type: 'gameSync', game: { ...safeGame, myHand: [...game.playerStates[tId].hand] } });
		}
	});
	renderGame({ ...safeGame, myHand: [...game.playerStates[myPeerId].hand] });
}

function handleClientData(data) {
	if (data.type === 'ping') {
		if (hostConnection && hostConnection.open) { hostConnection.send({ type: 'pong' }); }
	}
	else if (data.type === 'lobbySync') { players = data.players; renderLobby(); }
	else if (data.type === 'gameSync') {
		if ((game.status === 'waiting' || game.status === 'ended') && data.game.status === 'playing') switchScreen('screen-game', 'top-status-bar');
		game = { ...game, ...data.game }; renderGame(data.game);
	}
	else if (data.type === 'announce') { announce(data.message, data.assertive); logEvent(data.message); }
	else if (data.type === 'endGame') { showResult(data.winnerName, data.resultStats); }
	else if (data.type === 'playSound') { playSound(data.soundName); }
	else if (data.type === 'startAnim') { doStartAnimation(() => {}); }
}

function canPlayCard(card, activeSuit, topCard) {
    if (card.rank === '8') return true;
    if (card.suit === activeSuit) return true;
    if (card.rank === topCard.rank) return true;
    return false;
}

function renderGame(gState) {
	const turnPlayer = gState.players[gState.turnIndex];
	if (turnPlayer) {
		const headingEl = document.getElementById('current-turn-heading');
		headingEl.textContent = `รอบของ ${turnPlayer.name}`;
		headingEl.style.color = '#ffffff';
		if(turnPlayer.id === myPeerId) {
			headingEl.style.background = 'rgba(46, 125, 50, 0.85)';
			headingEl.style.borderColor = 'var(--primary)';
		} else {
			headingEl.style.background = 'rgba(15, 23, 42, 0.85)';
			headingEl.style.borderColor = 'rgba(255, 255, 255, 0.2)';
		}
	}

	const statusBar = document.getElementById('top-status-bar');
	statusBar.innerHTML = '';
	let ariaStatusBarText = ``;

	const mascots = ['🐱', '🐶', '🐰', '🦊', '🐼', '🐸'];
	gState.players.forEach((p, idx) => {
		const pState = gState.playerStates[p.id];
		const isTurn = idx === gState.turnIndex;
		ariaStatusBarText += `${p.name} มี ${pState.cardCount} ใบ. `;
		const charDiv = document.createElement('div');
		charDiv.className = `character-card ${isTurn ? 'is-turn' : ''}`; charDiv.setAttribute('aria-hidden', 'true');
		const mascot = mascots[idx % 6];
		charDiv.innerHTML = `<div class="char-cards-count">เหลือ ${pState.cardCount} ใบ</div><div class="char-name">${mascot} ${p.name}</div>`;
		statusBar.appendChild(charDiv);
	});
	statusBar.setAttribute('aria-label', ariaStatusBarText);

	document.getElementById('deck-count-visual').textContent = gState.deckCount;
	document.getElementById('deck-aria-label').textContent = `กองจั่วเหลือ ${gState.deckCount} ใบ`;

	const centerCont = document.getElementById('board-center-container');
	const centerAria = document.getElementById('board-center-aria');
	centerCont.innerHTML = '';
    
    if(gState.topCard) {
        centerCont.appendChild(renderCardHTML(gState.topCard, -1, false, true));
        centerAria.textContent = `กองทิ้ง: ${gState.topCard.name}`;
    }

	const myContainer = document.getElementById('my-cards-container');
	myContainer.innerHTML = '';
	const isMyTurn = (turnPlayer && turnPlayer.id === myPeerId && gState.status === 'playing');
	
	let canPlayAtLeastOne = false;
	gState.myHand.forEach((card, idx) => {
		let playable = false;
		if (isMyTurn) {
            playable = canPlayCard(card, gState.activeSuit, gState.topCard);
		}
		if (playable) canPlayAtLeastOne = true;
		myContainer.appendChild(renderCardHTML(card, idx, playable));
	});

	const btnDraw = document.getElementById('btn-draw');
	if (isMyTurn && !canPlayAtLeastOne && gState.deckCount > 0) {
		btnDraw.disabled = false;
		btnDraw.onclick = () => { btnDraw.disabled = true; sendAction('draw'); };
	} else {
		btnDraw.disabled = true;
	}
}

function sendAction(action, payload = null) {
	if (isHost) handlePlayerAction(myPeerId, action, payload);
	else if (hostConnection && hostConnection.open) hostConnection.send({ type: 'action', action, payload });
}

// --- HOST GAME LOGIC ---
let botTimer = null;

function handlePlayerAction(peerId, action, payload) {
	if (!isHost || game.status !== 'playing') return;
	const currentPlayer = players[game.turnIndex];

	if (action === 'draw' && peerId === currentPlayer.id) {
        // Draw 1 card and pass turn immediately
		if (game.deck.length === 0) return;
		const card = game.deck.pop();
		game.playerStates[peerId].hand.push(card);
		broadcastSound('jua');
		broadcastAnnounce(`${currentPlayer.name} จั่วไพ่แล้วจบตา`);
		broadcastGameState();
        
        setTimeout(() => advanceTurn(), 1500);
		return;
	}

	if (peerId !== currentPlayer.id || action !== 'play') return;
	
	const state = game.playerStates[peerId];
	if (state.isProcessing) return;
	clearTimeout(botTimer);
	state.isProcessing = true;

	const card = state.hand[payload.index];
	if (!card) { state.isProcessing = false; return; }

	state.hand.splice(payload.index, 1);
    game.discardPile.push(card);
    
    if (card.rank === '8') {
        game.activeSuit = payload.activeSuit;
        broadcastAnnounce(`${currentPlayer.name} ลงไพ่ 8 และเปลี่ยนดอกเป็น ${game.activeSuit}`);
    } else {
        game.activeSuit = card.suit;
        broadcastAnnounce(`${currentPlayer.name} ลง ${card.name}`);
    }

	broadcastSound('select');
	broadcastGameState();

	let delayBeforeEffect = 1500;
	if (state.hand.length === 1) {
		setTimeout(() => { broadcastAnnounce(`${currentPlayer.name} เหลือ 1 ใบ`); }, delayBeforeEffect);
		delayBeforeEffect += 1000;
	}

	if (state.hand.length === 0) { 
		setTimeout(() => { state.isProcessing = false; handleWin(peerId); }, delayBeforeEffect); 
		return; 
	}

	const finishPlay = () => {
		state.isProcessing = false;
        advanceTurn();
	};
    setTimeout(finishPlay, delayBeforeEffect);
}

function reshuffleDeck() {
    if(game.deck.length === 0 && game.discardPile.length > 1) {
        // Reshuffle discard pile to deck
        let top = game.discardPile.pop();
        game.deck = game.discardPile.sort(() => Math.random() - 0.5);
        game.discardPile = [top];
        broadcastAnnounce('กองจั่วหมด สับไพ่ใหม่เรียบร้อยแล้ว');
    }
}

function advanceTurn() {
    reshuffleDeck();
	game.turnIndex = (game.turnIndex + game.direction + players.length) % players.length;
	broadcastSound('turn');
	broadcastGameState();
	
	setTimeout(() => {
		broadcastTurnStart();
		processTurnLogic();
	}, 400);
}

async function runBotTurn(botPlayer, state) {
	await new Promise(r => setTimeout(r, 2000));
    const topCard = game.discardPile[game.discardPile.length - 1];

    let playableIndex = -1;
    for(let i = 0; i < state.hand.length; i++) {
        if (canPlayCard(state.hand[i], game.activeSuit, topCard)) {
            playableIndex = i;
            break;
        }
    }

	if (playableIndex !== -1) {
        const card = state.hand[playableIndex];
        let payload = { index: playableIndex };
        if (card.rank === '8') {
            // Bot chooses random suit when playing 8
            const suits = ['♥', '♦', '♣', '♠'];
            payload.activeSuit = suits[Math.floor(Math.random() * suits.length)];
        }
		handlePlayerAction(botPlayer.id, 'play', payload);
	} else {
        if(game.deck.length > 0) {
            handlePlayerAction(botPlayer.id, 'draw');
        } else {
            advanceTurn();
        }
	}
}

function processTurnLogic() {
	if (!isHost || game.status !== 'playing') return;
	const currentPlayer = players[game.turnIndex];
	const state = game.playerStates[currentPlayer.id];
	clearTimeout(botTimer);
	
	if (currentPlayer.isBot) {
		runBotTurn(currentPlayer, state);
	} else {
		botTimer = setTimeout(() => {
			if (game.status === 'playing' && players[game.turnIndex].id === currentPlayer.id) {
                // Auto-pass/draw logic timeout
                const topCard = game.discardPile[game.discardPile.length - 1];
                let hasPlayable = false;
                for(let i=0; i<state.hand.length; i++) {
                    if(canPlayCard(state.hand[i], game.activeSuit, topCard)) hasPlayable = true;
                }
                if (!hasPlayable && game.deck.length > 0) {
                    handlePlayerAction(currentPlayer.id, 'draw');
                }
			}
		}, 45000);
	}
}

function handleWin(winnerId) {
	game.status = 'ended';
	const winner = players.find(p => p.id === winnerId);
	
	let resultStats = [];
	players.forEach(p => {
        let cardsLeft = game.playerStates[p.id].hand.length;
		resultStats.push({ id: p.id, name: p.name, cardsLeft: cardsLeft, isWinner: p.id === winnerId });
	});

	broadcastGameState();
	let announceMsg = `การแข่งขันจบแล้ว ${winner.name} เป็นผู้ชนะ `;
	broadcastAnnounce(announceMsg, true);
	broadcastSound('win');

	connections.forEach(c => {
		if(c.open) {
			c.send({ type: 'endGame', winnerName: winner.name, resultStats: resultStats });
		}
	});
	showResult(winner.name, resultStats);
}

function showResult(winnerName, resultStats) {
	switchScreen('screen-result', 'title-result');
	stopBGM();
	let winnerStat = resultStats.find(r => r.isWinner) || resultStats[0];
	let losers = resultStats.filter(r => !r.isWinner);
	let html = `<div style="color: var(--focus-ring); margin-bottom: 10px;">👑 ${winnerName} เป็นผู้ชนะ (ไพ่หมดมือ)</div>`;
	losers.forEach(l => {
		html += `<div style="color: #ffcdd2; margin-bottom: 10px;">❌ ${l.name} เหลือไพ่ ${l.cardsLeft} ใบ</div>`;
	});
	document.getElementById('winner-text').innerHTML = html;
	document.getElementById('result-status-bar').innerHTML = '';
}