import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { getDatabase, ref, update, remove, onValue, get } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-database.js";

const firebaseConfig = {
	apiKey: "AIzaSyDvcdgsyT5sDdYTYKIqetzNL9Be-MFC0l4",
	authDomain: "xo-game-134ec.firebaseapp.com",
	databaseURL: "https://xo-game-134ec-default-rtdb.asia-southeast1.firebasedatabase.app",
	projectId: "xo-game-134ec",
	storageBucket: "xo-game-134ec.firebasestorage.app",
	messagingSenderId: "318375224157",
	appId: "1:318375224157:web:6a2bc6432e96e549b77eb4"
};
const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

// --- Audio System (Unchanged Timing & Logic) ---
const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
const soundBuffers = {};
const soundNames = ['1', 'select', 'start', 'bgm', 'jua', 'turn', 'uno', 'skip', 'reverse', 'draw2', 'draw4', 'll', 'ww', 'wl', 'win', 'abc', 'hit', 'sleep'];
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

function showAnimOverlay(text, isDomino = false) {
	if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
	const overlay = document.getElementById('anim-overlay');
	const textEl = document.getElementById('anim-text');
	textEl.textContent = text;
	textEl.className = isDomino ? 'anim-domino' : 'anim-text';
	overlay.style.display = 'flex';
	setTimeout(() => { overlay.style.display = 'none'; }, 1500);
}

function translateTile(val) {
	if(val === 'sleep') return 'หลับ';
	if(val === 'draw2') return '2+';
	if(val === 'draw3') return '3+';
	if(val === 'reverse') return 'ย้อนศร';
	return val;
}

// --- Visual FX Layer Functions (Non-blocking) ---
function getPipLayoutHTML(num) {
	if (typeof num !== 'number' || num === 0) return `<div class="pips-grid"></div>`;
	const dots = Array(9).fill(0).map((_, i) => `<div class="pip pip-id-${i}"></div>`);
	// Helper to hide specific pips based on number to create standard domino layout
	let hideIndices = [];
	if(num === 1) hideIndices = [0,1,2,3,5,6,7,8];
	else if(num === 2) hideIndices = [0,1,3,4,5,7,8]; // top-right, bottom-left (grid index 2, 6)
	else if(num === 3) hideIndices = [0,1,3,5,7,8]; // 2, 4, 6
	else if(num === 4) hideIndices = [1,3,4,5,7]; // 0, 2, 6, 8
	else if(num === 5) hideIndices = [1,3,5,7]; // 0, 2, 4, 6, 8
	else if(num === 6) hideIndices = [1,4,7]; // 0, 2, 3, 5, 6, 8
	else if(num === 7) hideIndices = [1,7]; // 0, 2, 3, 4, 5, 6, 8
	else if(num === 8) hideIndices = [4]; // 8 pips, hide center
	else if(num === 9) hideIndices = []; // all 9
	
	return `<div class="pips-grid pips-${num}">` + dots.map((html, i) => {
		if (hideIndices.includes(i)) return `<div class="pip hidden"></div>`;
		return html;
	}).join('') + `</div>`;
}

const fxLayer = document.getElementById('visual-fx-layer');

function fxPlayTile(tile, side, sourceRect) {
	if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
	const targetId = side === 'left' ? 'board-left-container' : 'board-right-container';
	const targetEl = document.getElementById(targetId);
	if(!targetEl || !sourceRect) return;
	const destRect = targetEl.getBoundingClientRect();
	
	const fxEl = document.createElement('div');
	fxEl.className = 'fx-tile';
	fxEl.innerHTML = renderTileHTML(tile, -1, false, false, true).outerHTML; // clone visual
	
	const startX = sourceRect.left + sourceRect.width/2 - 25; // approx center
	const startY = sourceRect.top + sourceRect.height/2 - 40;
	const endX = destRect.left + destRect.width/2 - 25;
	const endY = destRect.top + destRect.height/2 - 40;
	
	fxEl.style.setProperty('--startX', `${startX}px`);
	fxEl.style.setProperty('--startY', `${startY}px`);
	fxEl.style.setProperty('--endX', `${endX}px`);
	fxEl.style.setProperty('--endY', `${endY}px`);
	fxEl.style.setProperty('--rot', `${side==='left'? -15 : 15}deg`);
	fxEl.style.animation = `fxFlyToBoard 0.6s cubic-bezier(0.25, 1, 0.5, 1) forwards`;
	
	fxLayer.appendChild(fxEl);
	setTimeout(() => fxEl.remove(), 600);
}

function fxDrawTile(targetPeerId) {
	if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
	const deckEl = document.getElementById('ui-deck-pile');
	let targetEl = document.getElementById(`char-${targetPeerId}`);
	if(!targetEl) targetEl = document.getElementById('my-cards-container'); // fallback
	
	if(!deckEl || !targetEl) return;
	const deckRect = deckEl.getBoundingClientRect();
	const targetRect = targetEl.getBoundingClientRect();

	const fxEl = document.createElement('div');
	fxEl.className = 'fx-tile domino-tile';
	fxEl.style.background = 'linear-gradient(145deg, #1f2937, #111827)'; // back of tile matching dark deck
	
	const startX = deckRect.left; const startY = deckRect.top;
	const endX = targetRect.left + targetRect.width/2; const endY = targetRect.top + targetRect.height/2;

	fxEl.style.setProperty('--startX', `${startX}px`); fxEl.style.setProperty('--startY', `${startY}px`);
	fxEl.style.setProperty('--endX', `${endX}px`); fxEl.style.setProperty('--endY', `${endY}px`);
	fxEl.style.animation = `fxFlyFromDeck 0.5s ease-out forwards`;
	
	fxLayer.appendChild(fxEl);
	setTimeout(() => fxEl.remove(), 500);
}

function fxSpecialEffect(type, targetPeerId) {
	if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
	const charEl = document.getElementById(`char-${targetPeerId}`);
	if(!charEl) return;
	
	if(type === 'sleep') {
		charEl.style.setProperty('--glowColor', 'var(--sp-sleep)');
		charEl.classList.add('fx-glow');
	} else if(['draw2', 'draw3'].includes(type)) {
		charEl.style.setProperty('--glowColor', type === 'draw3' ? 'var(--sp-draw3)' : 'var(--sp-draw2)');
		charEl.classList.add('fx-shake', 'fx-glow');
	} else if(type === 'reverse') {
		const dirEl = document.querySelector('.direction-indicator');
		if(dirEl) { dirEl.style.transform += ' scale(1.5)'; setTimeout(()=>dirEl.style.transform = dirEl.style.transform.replace(' scale(1.5)',''), 400); }
	}
	setTimeout(() => { charEl.classList.remove('fx-shake', 'fx-glow'); }, 1500);
}

function fxCelebrate() {
	if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
	for(let i=0; i<50; i++) {
		const conf = document.createElement('div');
		conf.style.position = 'absolute';
		conf.style.left = Math.random() * 100 + 'vw';
		conf.style.top = '-10px';
		conf.style.width = Math.random() * 10 + 5 + 'px';
		conf.style.height = Math.random() * 10 + 5 + 'px';
		conf.style.backgroundColor = ['#ffeb3b', '#4CAF50', '#f44336', '#0ea5e9'][Math.floor(Math.random()*4)];
		conf.style.animation = `fxConfettiFall ${Math.random()*2 + 2}s linear forwards`;
		fxLayer.appendChild(conf);
		setTimeout(() => conf.remove(), 4000);
	}
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

const MAX_PLAYERS = 6;
const BOT_NAMES = [
	'บอทซุปเปอร์แมน', 'บอทแมงมุม', 'บอทมาริโอ้', 'บอทนารูโตะ', 'บอทโซนิค', 
	'บอทปิกาจู', 'บอทเบ็นเท็น', 'บอทโดราเอมอน', 'บอทลูฟี่', 'บอทคิริโตะ',
	'บอทไอรอนแมน', 'บอทแบทแมน', 'บอทกัปตัน', 'บอทชินจัง', 'บอทโคนัน', 
	'บอทอุลตร้าแมน', 'บอทก็อดซิลล่า', 'บอทเซเลอร์มูน', 'บอทซากุระ', 'บอทดราก้อนบอล'
];

let players = [];
let game = {
	status: 'waiting',
	deck: [],
	leftEnd: null,
	rightEnd: null,
	turnIndex: 0,
	direction: 1,
	playerStates: {}, 
	consecutivePasses: 0,
	matchOver: false,
	pendingSkips: 0,
	drawCountThisTurn: 0,
	doubleChainSide: null
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

// --- UI Navigation ---
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
	if (e.key === 'Enter') {
		e.preventDefault();
		if (nameInput.value.trim().length > 0) {
			btnConfirmName.click();
		}
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
const roomsRef = ref(db, 'domino_rooms');
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
	if (emptyMsg) {
		emptyMsg.remove();
	}

	const existingItems = Array.from(list.querySelectorAll('li[data-room-id]'));
	const activeRoomIds = activeRooms.map(([id]) => id);

	// Remove rooms that are no longer active
	existingItems.forEach(li => {
		const roomId = li.getAttribute('data-room-id');
		if (!activeRoomIds.includes(roomId)) {
			li.remove();
		}
	});

	// Update or create rooms
	activeRooms.forEach(([id, room]) => {
		let li = list.querySelector(`li[data-room-id="${id}"]`);
		if (li) {
			// Update existing button to maintain accessibility focus
			const btn = li.querySelector('button');
			if (btn) {
				btn.innerHTML = `<span>ห้อง: ${id}</span> <span>${room.currentPlayers}/${MAX_PLAYERS} คน</span>`;
				btn.setAttribute('aria-label', `เข้าร่วมห้อง ${id} มีผู้เล่น ${room.currentPlayers} จาก ${MAX_PLAYERS} คน`);
			}
		} else {
			// Create new room item
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
		const metaSnapshot = await get(ref(db, 'domino_metadata/last_room_id'));
		if (metaSnapshot.exists()) {
			nextIdNum = metaSnapshot.val() + 1;
			if (nextIdNum > 99999) nextIdNum = 1;
		}
		await update(ref(db, 'domino_metadata'), { last_room_id: nextIdNum });
		
		currentRoomId = "domino" + String(nextIdNum).padStart(5, '0');
		isHost = true;
		players = [{ id: myPeerId, name: myName, isBot: false }];
		
		clearInterval(heartbeatInterval);
		heartbeatInterval = setInterval(hostHeartbeatCheck, 3000);

		await update(ref(db, `domino_rooms/${currentRoomId}`), { hostPeerId: myPeerId, status: 'waiting', currentPlayers: 1, lastActive: Date.now() });

		setInterval(() => { if (isHost && currentRoomId && game.status === 'waiting') update(ref(db, `domino_rooms/${currentRoomId}`), { lastActive: Date.now() }); }, 3000);
		enterLobby();
		document.getElementById('lobby-ready-msg').style.display = 'block';
		announce('ห้องโดมิโน่พร้อมแล้ว รอเพื่อนหรือเพิ่มบอทได้ทันที');
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
	if (isHost && currentRoomId) { remove(ref(db, `domino_rooms/${currentRoomId}`)); connections.forEach(c => c.close()); }
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
document.getElementById('lobby-ready-msg').innerText = 'เล่นได้สูงสุดที่ 6 คนต่อห้อง เจ้าของห้องพร้อมก็กดเริ่มเกมได้นะ เพื่อสนุกกัน';document.getElementById('lobby-ready-msg').style.display = 'block';
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
		broadcastAnnounce(`${p.name} หลุดการเชื่อมต่อ เปลี่ยนเป็นบอทแล้ว`, true);
		syncLobby(); 
		if (game.status === 'playing') {
			broadcastGameState();
			if (players[game.turnIndex].id === peerId) {
				processTurnLogic();
			}
		}
	}
}

function hostHeartbeatCheck() {
	if (!isHost) return;
	const now = Date.now();
	connections.forEach(conn => {
		if (!conn.lastPing) conn.lastPing = now;
		
		if (conn.open) {
			conn.send({ type: 'ping' });
		}
		
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
			update(ref(db, `domino_rooms/${currentRoomId}`), { currentPlayers: players.length });
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
		broadcastAnnounce(`เพิ่ม ${botName} เข้าห้องแล้ว`, true);
		if (currentRoomId) update(ref(db, `domino_rooms/${currentRoomId}`), { currentPlayers: players.length });
	}
};
document.getElementById('btn-remove-bot').onclick = () => {
	const botIdx = players.slice().reverse().findIndex(p => p.isBot);
	if (botIdx !== -1) {
		const botToRemove = players[players.length - 1 - botIdx];
		players.splice(players.length - 1 - botIdx, 1);
		syncLobby(); broadcastSound('select');
		broadcastAnnounce(`ลด ${botToRemove.name} ออกจากห้องแล้ว`, true);
		if (currentRoomId) update(ref(db, `domino_rooms/${currentRoomId}`), { currentPlayers: players.length });
	}
};

function syncLobby() {
	if (!isHost) return;
	renderLobby();
	const safePlayers = players.map(p => ({ id: p.id, name: p.name, isBot: p.isBot }));
	connections.forEach(c => { if(c.open) c.send({ type: 'lobbySync', players: safePlayers }); });
}

// --- DOMINO LOGIC ---
function generateDeck() {
	const deck = [];
	let idCount = 0;
	for (let i = 0; i <= 9; i++) {
		for (let j = i; j <= 9; j++) {
			deck.push({ id: idCount++, left: i, right: j, type: 'regular', name: `${i} กับ ${j}` });
		}
	}
	const specials = [
		{ val: 0, effect: 'sleep', name: 'หลับ' }, { val: 2, effect: 'sleep', name: 'หลับ' }, { val: 4, effect: 'sleep', name: 'หลับ' }, { val: 6, effect: 'sleep', name: 'หลับ' }, { val: 8, effect: 'sleep', name: 'หลับ' },
		{ val: 1, effect: 'draw2', name: '2+' }, { val: 3, effect: 'draw2', name: '2+' }, { val: 5, effect: 'draw2', name: '2+' }, { val: 7, effect: 'draw2', name: '2+' }, { val: 9, effect: 'draw2', name: '2+' },
		{ val: 0, effect: 'draw3', name: '3+' }, { val: 2, effect: 'draw3', name: '3+' }, { val: 4, effect: 'draw3', name: '3+' }, { val: 6, effect: 'draw3', name: '3+' }, { val: 8, effect: 'draw3', name: '3+' },
		{ val: 1, effect: 'reverse', name: 'ย้อนศร' }, { val: 3, effect: 'reverse', name: 'ย้อนศร' }, { val: 5, effect: 'reverse', name: 'ย้อนศร' }, { val: 7, effect: 'reverse', name: 'ย้อนศร' }, { val: 9, effect: 'reverse', name: 'ย้อนศร' }
	];
	specials.forEach(sp => {
		deck.push({ id: idCount++, left: sp.val, right: sp.effect, type: sp.effect, name: `${sp.val} กับ ${sp.name}` });
	});
	return deck.sort(() => Math.random() - 0.5);
}

function getTileScore(tile) {
	if(tile.type === 'regular') return tile.left + tile.right;
	return 20; 
}

function getTileAria(tile) { return `${tile.name}`; }

function renderTileHTML(tile, index = -1, playableLeft = false, playableRight = false, forceVisualOnly = false) {
	const isSpecial = tile.type !== 'regular';
	const isDouble = tile.left === tile.right && tile.type === 'regular';
	const btn = document.createElement((index >= 0 && !forceVisualOnly) ? 'button' : 'div');
	
	let spClass = '';
	if(tile.type === 'sleep') spClass = 'sp-sleep';
	else if(tile.type === 'draw2') spClass = 'sp-draw2';
	else if(tile.type === 'draw3') spClass = 'sp-draw3';
	else if(tile.type === 'reverse') spClass = 'sp-reverse';
	
	btn.className = `domino-tile ${isDouble ? '' : 'horizontal'} ${isSpecial ? 'special-tile ' + spClass : ''} ${(index >= 0 && !playableLeft && !playableRight && !forceVisualOnly) ? 'disabled' : ''}`;
	
	const formatHalf = (val) => {
		if(val === 'sleep') return `<div class="domino-half special-text">หลับ<br>💤</div>`;
		if(val === 'draw2') return `<div class="domino-half special-text">2+<br>⚡</div>`;
		if(val === 'draw3') return `<div class="domino-half special-text">3+<br>💥</div>`;
		if(val === 'reverse') return `<div class="domino-half special-text">ย้อนศร<br>🔄</div>`;
		return `<div class="domino-half">${getPipLayoutHTML(val)}</div>`;
	};

	btn.innerHTML = `${formatHalf(tile.left)}${formatHalf(tile.right)}`;
	
	if (index >= 0 && !forceVisualOnly) {
		btn.setAttribute('aria-label', `${getTileAria(tile)} ${playableLeft || playableRight ? 'ลงได้' : 'ลงไม่ได้'}`);
		if (playableLeft || playableRight) {
			btn.onclick = (e) => {
				const rect = btn.getBoundingClientRect();
				handleTileClick(index, tile, playableLeft, playableRight, rect);
			};
		} else {
			btn.setAttribute('aria-disabled', 'true');
		}
	}
	return btn;
}

document.getElementById('btn-start-game').onclick = () => {
	document.getElementById('btn-start-game').disabled = true;
	if (isHost && players.length >= 2) {
		remove(ref(db, `domino_rooms/${currentRoomId}`));
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
		{ time: 200, text: 'DOMINO' }, { time: 1200, text: '3' }, { time: 2200, text: '2' }, { time: 3200, text: '1' }, { time: 4200, text: 'Enjoy' }
	];
	steps.forEach(step => {
		setTimeout(() => { animDiv.textContent = step.text; announce(step.text, true); }, step.time);
	});
	setTimeout(() => { animDiv.remove(); if(callback) callback(); }, 5000);
}

function broadcastTurnStart() {
	if (!isHost) return;
	const currentPlayer = players[game.turnIndex];
	let dirStr = game.direction === 1 ? 'ตามเข็มนาฬิกา' : 'ทวนเข็มนาฬิกา';
	let msg = `ถึงรอบของ ${currentPlayer.name} ขณะนี้เล่น${dirStr} `;
	if (game.leftEnd === null) {
		msg += `กระดานยังว่าง คุณสามารถเริ่มวางได้`;
	} else {
		msg += `ปลายซ้ายคือ ${translateTile(game.leftEnd)} ปลายขวาคือ ${translateTile(game.rightEnd)}`;
	}
	if (game.doubleChainSide) {
		msg += ` (มีการล็อกฝั่งเล่นต่อที่ปลาย${game.doubleChainSide === 'left' ? 'ซ้าย' : 'ขวา'})`;
	}
	broadcastAnnounce(msg);
}

function initGame() {
	game.status = 'playing'; game.matchOver = false;
	game.deck = generateDeck();
	game.leftEnd = null; game.rightEnd = null;
	game.turnIndex = 0;
	game.direction = 1; game.consecutivePasses = 0;
	game.pendingSkips = 0;
	game.drawCountThisTurn = 0;
	game.doubleChainSide = null;
	
	let highestDoubleVal = -1;
	let starterIdx = 0;

	players.forEach((p, i) => {
		game.playerStates[p.id] = { hand: [], declaredDomino: false };
		for(let k=0; k<7; k++) game.playerStates[p.id].hand.push(game.deck.pop());
		
		game.playerStates[p.id].hand.forEach(t => {
			if(t.type === 'regular' && t.left === t.right && t.left > highestDoubleVal) {
				highestDoubleVal = t.left; starterIdx = i;
			}
		});
	});

	game.turnIndex = starterIdx;
	switchScreen('screen-game', 'top-status-bar');
	broadcastGameState();
	
	setTimeout(() => {
		broadcastTurnStart();
		broadcastSound('turn');
		const firstPlayer = players[game.turnIndex];
		if (firstPlayer.id === myPeerId && !firstPlayer.isBot) {
			playSound('abc');
		} else if (!firstPlayer.isBot && firstPlayer.connection && firstPlayer.connection.open) {
			firstPlayer.connection.send({ type: 'playAbc' });
		}
		processTurnLogic();
	}, 1000);
}

function broadcastGameState() {
	if (!isHost) return;
	const safePlayers = players.map(p => ({ id: p.id, name: p.name, isBot: p.isBot }));
	const safeGame = {
		status: game.status,
		deckCount: game.deck.length,
		leftEnd: game.leftEnd,
		rightEnd: game.rightEnd,
		turnIndex: game.turnIndex,
		direction: game.direction,
		drawCountThisTurn: game.drawCountThisTurn,
		doubleChainSide: game.doubleChainSide,
		players: safePlayers,
		playerStates: {}
	};
	players.forEach(p => {
		safeGame.playerStates[p.id] = { cardCount: game.playerStates[p.id].hand.length, declaredDomino: game.playerStates[p.id].declaredDomino };
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
		if (hostConnection && hostConnection.open) {
			hostConnection.send({ type: 'pong' });
		}
	}
	else if (data.type === 'lobbySync') { players = data.players; renderLobby(); }
	else if (data.type === 'gameSync') {
		if ((game.status === 'waiting' || game.status === 'ended') && data.game.status === 'playing') switchScreen('screen-game', 'top-status-bar');
		game = { ...game, ...data.game }; renderGame(data.game);
	}
	else if (data.type === 'announce') { announce(data.message, data.assertive); logEvent(data.message); }
	else if (data.type === 'animDomino') { showAnimOverlay('DOMINO!', true); }
	else if (data.type === 'endGame') { showResult(data.winnerName, data.resultStats); }
	else if (data.type === 'playSound') { playSound(data.soundName); }
	else if (data.type === 'startAnim') { doStartAnimation(() => {}); }
	else if (data.type === 'playAbc') { playSound('abc'); }
	else if (data.type === 'fxDraw') { fxDrawTile(data.peerId); }
	else if (data.type === 'fxSpecial') { fxSpecialEffect(data.spType, data.targetId); }
	else if (data.type === 'fxCelebrate') { fxCelebrate(); }
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
	let ariaStatusBarText = `ทิศทาง ${gState.direction === 1 ? 'ตามเข็ม' : 'ทวนเข็ม'}, `;
	const dirDiv = document.createElement('div'); dirDiv.className = 'direction-indicator'; dirDiv.setAttribute('aria-hidden', 'true');
	dirDiv.style.transform = gState.direction === 1 ? 'rotate(0deg)' : 'rotate(180deg)'; dirDiv.textContent = '➡️';
	statusBar.appendChild(dirDiv);

	const mascots = ['🐱', '🐶', '🐰', '🦊', '🐼', '🐸'];
	gState.players.forEach((p, idx) => {
		const pState = gState.playerStates[p.id];
		const isTurn = idx === gState.turnIndex;
		ariaStatusBarText += `${p.name} มี ${pState.cardCount} ตัว. `;
		const charDiv = document.createElement('div');
		charDiv.id = `char-${p.id}`;
		charDiv.className = `character-card ${isTurn ? 'is-turn' : ''}`; charDiv.setAttribute('aria-hidden', 'true');
		const mascot = mascots[idx % 6];
		charDiv.innerHTML = `<div style="display: none;"><div class="char-cards-count">ไพ่ ${pState.cardCount}</div><div class="char-name">${p.isBot?'🤖 ':'👤 '}${p.name}</div></div><div class="char-cards-count">เหลือ ${pState.cardCount} ตัว</div><div class="char-name">${mascot} ${p.name}</div>`;
		statusBar.appendChild(charDiv);
	});
	statusBar.setAttribute('aria-label', ariaStatusBarText);

	document.getElementById('deck-count-visual').textContent = gState.deckCount;
	document.getElementById('deck-aria-label').textContent = `กองจั่วเหลือ ${gState.deckCount} ตัว`;

	// Render Board Ends visually
	const lCont = document.getElementById('board-left-container');
	const rCont = document.getElementById('board-right-container');
	const lAria = document.getElementById('board-left-aria');
	const rAria = document.getElementById('board-right-aria');
	
	// Cleanup visual nodes but keep ARIA labels
	Array.from(lCont.children).forEach(c => { if(c.className !== 'board-label') c.remove(); });
	Array.from(rCont.children).forEach(c => { if(c.className !== 'board-label') c.remove(); });

	if (gState.leftEnd === null) {
		lAria.textContent = 'กระดานว่างเปล่า'; rAria.textContent = 'กระดานว่างเปล่า';
	} else {
		const renderEndVisual = (val) => {
			const d = document.createElement('div');
			d.className = 'domino-end-tile domino-half';
			d.setAttribute('aria-hidden', 'true');
			if (val === 'sleep') {
				d.classList.add('special-tile', 'sp-sleep');
				d.innerHTML = `<div class="special-text">หลับ<br>💤</div>`;
			} else if (val === 'draw2') {
				d.classList.add('special-tile', 'sp-draw2');
				d.innerHTML = `<div class="special-text">2+<br>⚡</div>`;
			} else if (val === 'draw3') {
				d.classList.add('special-tile', 'sp-draw3');
				d.innerHTML = `<div class="special-text">3+<br>💥</div>`;
			} else if (val === 'reverse') {
				d.classList.add('special-tile', 'sp-reverse');
				d.innerHTML = `<div class="special-text">ย้อนศร<br>🔄</div>`;
			} else {
				d.innerHTML = getPipLayoutHTML(val);
			}
			return d;
		};
		lCont.appendChild(renderEndVisual(gState.leftEnd));
		rCont.appendChild(renderEndVisual(gState.rightEnd));
		lAria.textContent = `ปลายซ้าย ${translateTile(gState.leftEnd)}`;
		rAria.textContent = `ปลายขวา ${translateTile(gState.rightEnd)}`;
	}

	// Render Hand
	const myContainer = document.getElementById('my-cards-container');
	myContainer.innerHTML = '';
	
	const isMyTurn = (turnPlayer && turnPlayer.id === myPeerId && gState.status === 'playing');
	
	let canPlayAtLeastOne = false;
	gState.myHand.forEach((tile, idx) => {
		let playableL = false, playableR = false;
		if (isMyTurn) {
			if (gState.leftEnd === null) { playableL = true; playableR = true; }
			else {
				if (!gState.doubleChainSide || gState.doubleChainSide === 'left') {
					if (tile.left === gState.leftEnd || tile.right === gState.leftEnd) playableL = true;
				}
				if (!gState.doubleChainSide || gState.doubleChainSide === 'right') {
					if (tile.left === gState.rightEnd || tile.right === gState.rightEnd) playableR = true;
				}
			}
		}
		if (playableL || playableR) canPlayAtLeastOne = true;
		myContainer.appendChild(renderTileHTML(tile, idx, playableL, playableR));
	});

	// Draw Button
	const btnDraw = document.getElementById('btn-draw');
	if (isMyTurn && gState.drawCountThisTurn < 3 && !canPlayAtLeastOne && gState.deckCount > 0) {
		btnDraw.disabled = false;
		btnDraw.onclick = () => { btnDraw.disabled = true; sendAction('draw'); };
	} else {
		btnDraw.disabled = true;
	}

	// DOMINO Button
	const btnDomino = document.getElementById('btn-domino');
	const myState = gState.playerStates[myPeerId];
	if (isMyTurn && gState.myHand.length === 2 && myState && !myState.declaredDomino && canPlayAtLeastOne) {
		btnDomino.disabled = false;
		btnDomino.onclick = () => { btnDomino.disabled = true; sendAction('announce_domino'); };
	} else { btnDomino.disabled = true; }
}

let pendingPlayTile = null;
window.handleTileClick = function(index, tile, playableLeft, playableRight, sourceRect) {
	if (playableLeft && playableRight && game.leftEnd !== null && game.leftEnd !== game.rightEnd) {
		pendingPlayTile = { index, tile, sourceRect };
		document.getElementById('side-picker-modal').style.display = 'flex';
		document.getElementById('btn-play-left').focus();
	} else if (playableLeft) { 
		fxPlayTile(tile, 'left', sourceRect);
		sendAction('play', { index, side: 'left' }); 
	} else if (playableRight) { 
		fxPlayTile(tile, 'right', sourceRect);
		sendAction('play', { index, side: 'right' }); 
	}
}

document.getElementById('btn-play-left').onclick = () => { document.getElementById('side-picker-modal').style.display = 'none'; fxPlayTile(pendingPlayTile.tile, 'left', pendingPlayTile.sourceRect); sendAction('play', { index: pendingPlayTile.index, side: 'left' }); };
document.getElementById('btn-play-right').onclick = () => { document.getElementById('side-picker-modal').style.display = 'none'; fxPlayTile(pendingPlayTile.tile, 'right', pendingPlayTile.sourceRect); sendAction('play', { index: pendingPlayTile.index, side: 'right' }); };
document.getElementById('btn-cancel-play').onclick = () => { document.getElementById('side-picker-modal').style.display = 'none'; pendingPlayTile = null; };

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
		if (game.drawCountThisTurn >= 3 || game.deck.length === 0) return;
		
		const tile = game.deck.pop();
		game.playerStates[peerId].hand.push(tile);
		game.playerStates[peerId].declaredDomino = false;
		game.drawCountThisTurn++;
		
		broadcastSound('jua');
		fxDrawTile(peerId);
		connections.forEach(c => { if(c.open) c.send({ type: 'fxDraw', peerId: peerId }); });

		players.forEach(p => {
			if (p.id === peerId) {
				const msg = `คุณจั่วได้ ${translateTile(tile.left)} กับ ${translateTile(tile.right)}`;
				if (p.connection && p.connection.open) p.connection.send({ type: 'announce', message: msg });
				else if (p.id === myPeerId) announce(msg);
			} else {
				const msg = `${currentPlayer.name} จั่ว 1 ตัว`;
				if (p.connection && p.connection.open) p.connection.send({ type: 'announce', message: msg });
				else if (p.id === myPeerId) announce(msg);
			}
		});

		broadcastGameState();

		if (game.drawCountThisTurn === 3) {
			if (!canPlayAny(game.playerStates[peerId].hand, game.doubleChainSide)) {
				setTimeout(() => {
					broadcastAnnounce(`${currentPlayer.name} จั่วครบ 3 ตัวแล้ว แต่ยังไม่มีตัวที่วางได้ จบตา`);
					game.playerStates[peerId].declaredDomino = false;
					game.consecutivePasses++;
					if (game.consecutivePasses >= players.length) { handleStuckGame(); return; }
					setTimeout(() => advanceTurn(), 2500);
				}, 1000);
			}
		}
		return;
	}

	if (action === 'announce_domino' && peerId === currentPlayer.id) {
		const state = game.playerStates[peerId];
		if (state.hand.length === 2 && !state.declaredDomino) {
			state.declaredDomino = true;
			broadcastAnnounce(`${currentPlayer.name} ประกาศ DOMINO! เหลือ 2 ตัว`);
			connections.forEach(c => { if(c.open) c.send({ type: 'animDomino' }); });
			showAnimOverlay('DOMINO!', true); 
			broadcastSound('uno');
			broadcastGameState();
		}
		return;
	}

	if (peerId !== currentPlayer.id || action !== 'play') return;
	
	const state = game.playerStates[peerId];
	if (state.isProcessing) return;
	
	if (game.doubleChainSide && game.doubleChainSide !== payload.side) return;

	clearTimeout(botTimer);
	state.isProcessing = true;

	const tile = state.hand[payload.index];
	if (!tile) { state.isProcessing = false; return; }

	// Visual Drop for others/bots
	if (peerId !== myPeerId) {
		const charEl = document.getElementById(`char-${peerId}`);
		if (charEl) fxPlayTile(tile, payload.side, charEl.getBoundingClientRect());
	}

	// Penalty Check
	if (state.hand.length === 2 && !state.declaredDomino) {
		broadcastAnnounce(`${currentPlayer.name} ลืมกด DOMINO! ถูกลงโทษ จั่ว 2 ตัว`);
		broadcastSound('wl');
		drawTiles(peerId, 2);
		state.declaredDomino = false;
	}

	state.hand.splice(payload.index, 1);
	let newEnd = null;
	let playedSideStr = '';
	let playedSide = null;

	if (game.leftEnd === null) { 
		game.leftEnd = tile.left; game.rightEnd = tile.right; 
		playedSideStr = 'เป็นตัวแรก';
	}
	else if (payload.side === 'left') {
		newEnd = (tile.left === game.leftEnd) ? tile.right : tile.left;
		game.leftEnd = newEnd;
		playedSideStr = 'ที่ปลายซ้าย';
		playedSide = 'left';
	} else {
		newEnd = (tile.left === game.rightEnd) ? tile.right : tile.left;
		game.rightEnd = newEnd;
		playedSideStr = 'ที่ปลายขวา';
		playedSide = 'right';
	}

	let isDouble = (tile.type === 'regular' && tile.left === tile.right);
	if (isDouble && playedSide) {
		game.doubleChainSide = playedSide;
	}

	let tileLeftText = translateTile(tile.left);
	let tileRightText = translateTile(tile.right);
	let actionStr = '';
	
	if (tile.type === 'regular') {
		actionStr = `${currentPlayer.name} วาง ${tileLeftText} กับ ${tileRightText} ${playedSideStr}`;
	} else {
		actionStr = `${currentPlayer.name} วาง ${translateTile(tile.type)} ${playedSideStr}`;
	}
	
	broadcastAnnounce(actionStr);
	broadcastSound('select');

	game.consecutivePasses = 0;
	broadcastGameState(); // อัปเดตกระดานให้ทุกคนเห็นทันที

	let delayBeforeEffect = 2000;

	if (state.hand.length === 1) {
		setTimeout(() => { broadcastAnnounce(`${currentPlayer.name} เหลือ 1 ตัว`); }, delayBeforeEffect);
		delayBeforeEffect += 1500;
	}

	if (state.hand.length === 0) { 
		setTimeout(() => { state.isProcessing = false; handleWin(peerId); }, delayBeforeEffect); 
		return; 
	}

	const finishPlay = () => {
		state.isProcessing = false;
		if (game.status !== 'playing' || players[game.turnIndex].id !== peerId) return;
		
		const st = game.playerStates[peerId];
		if (isDouble && st && st.hand.length > 0 && canPlayAny(st.hand, game.doubleChainSide)) {
			broadcastAnnounce(`${currentPlayer.name} วาง Double สามารถเล่นต่อเนื่องได้ที่ฝั่งเดิม`);
			broadcastGameState();
			processTurnLogic(); 
		} else {
			advanceTurn();
		}
	};

	// Apply special effect
	if (tile.type !== 'regular') {
		setTimeout(() => {
			if (tile.type === 'sleep') {
				const target = players[(game.turnIndex + game.direction + players.length) % players.length];
				broadcastAnnounce(`${target.name} ถูกข้ามตา`); 
				broadcastSound('sleep');
				game.pendingSkips = (game.pendingSkips || 0) + 1;
				fxSpecialEffect('sleep', target.id);
				connections.forEach(c => { if(c.open) c.send({ type: 'fxSpecial', spType: 'sleep', targetId: target.id }); });
			} else if (tile.type === 'reverse') {
				game.direction *= -1; 
				broadcastAnnounce(`ทิศทางการเล่นเปลี่ยนเป็น${game.direction === 1 ? 'ตามเข็มนาฬิกา' : 'ทวนเข็มนาฬิกา'}`); 
				broadcastSound('reverse');
				fxSpecialEffect('reverse', currentPlayer.id);
				connections.forEach(c => { if(c.open) c.send({ type: 'fxSpecial', spType: 'reverse', targetId: currentPlayer.id }); });
			} else if (tile.type === 'draw2') {
				const target = players[(game.turnIndex + game.direction + players.length) % players.length];
				broadcastAnnounce(`${target.name}โดน 2+ ต้องจั่วเพิ่ม 2 ตัวและถูกข้ามตา`); 
				broadcastSound('draw2');
				broadcastSound('hit');
				drawTiles(target.id, 2);
				game.pendingSkips = (game.pendingSkips || 0) + 1;
				fxSpecialEffect('draw2', target.id);
				connections.forEach(c => { if(c.open) c.send({ type: 'fxSpecial', spType: 'draw2', targetId: target.id }); });
			} else if (tile.type === 'draw3') {
				const target = players[(game.turnIndex + game.direction + players.length) % players.length];
				broadcastAnnounce(`${target.name}โดน 3+ ต้องจั่วเพิ่ม 3 ตัวและถูกข้ามตา`); 
				broadcastSound('draw4');
				broadcastSound('hit');
				drawTiles(target.id, 3);
				game.pendingSkips = (game.pendingSkips || 0) + 1;
				fxSpecialEffect('draw3', target.id);
				connections.forEach(c => { if(c.open) c.send({ type: 'fxSpecial', spType: 'draw3', targetId: target.id }); });
			}
		}, delayBeforeEffect);
		
		setTimeout(finishPlay, delayBeforeEffect + 2500); 
	} else {
		setTimeout(finishPlay, delayBeforeEffect);
	}
}

function drawTiles(playerId, count) {
	const st = game.playerStates[playerId];
	for(let i=0; i<count; i++) { if(game.deck.length > 0) st.hand.push(game.deck.pop()); }
	st.declaredDomino = false;
}

function canPlayAny(hand, lockedSide = null) {
	if (game.leftEnd === null) return hand.length > 0;
	if (lockedSide === 'left') {
		return hand.some(t => t.left === game.leftEnd || t.right === game.leftEnd);
	}
	if (lockedSide === 'right') {
		return hand.some(t => t.left === game.rightEnd || t.right === game.rightEnd);
	}
	return hand.some(t => t.left === game.leftEnd || t.right === game.leftEnd || t.left === game.rightEnd || t.right === game.rightEnd);
}

function advanceTurn() {
	let steps = 1 + (game.pendingSkips || 0);
	game.pendingSkips = 0;
	game.drawCountThisTurn = 0;
	game.doubleChainSide = null;
	game.turnIndex = (game.turnIndex + (game.direction * steps) + (players.length * 10)) % players.length;
	
	const nextPlayer = players[game.turnIndex];
	broadcastSound('turn');
	if (nextPlayer.id === myPeerId && !nextPlayer.isBot) {
		playSound('abc');
	} else if (!nextPlayer.isBot && nextPlayer.connection && nextPlayer.connection.open) {
		nextPlayer.connection.send({ type: 'playAbc' });
	}
	
	broadcastGameState();
	
	setTimeout(() => {
		broadcastTurnStart();
		processTurnLogic();
	}, 400);
}

async function runBotTurn(botPlayer, state) {
	await new Promise(r => setTimeout(r, 2000)); // wait for turn announcement

	while (game.drawCountThisTurn < 3 && !canPlayAny(state.hand, game.doubleChainSide) && game.deck.length > 0) {
		const tile = game.deck.pop();
		state.hand.push(tile);
		state.declaredDomino = false;
		game.drawCountThisTurn++;
		broadcastSound('jua');
		fxDrawTile(botPlayer.id);
		connections.forEach(c => { if(c.open) c.send({ type: 'fxDraw', peerId: botPlayer.id }); });
		broadcastAnnounce(`${botPlayer.name} จั่ว 1 ตัว`);
		broadcastGameState();
		await new Promise(r => setTimeout(r, 1500));
	}

	if (game.status !== 'playing') return;

	if (state.hand.length === 2 && canPlayAny(state.hand, game.doubleChainSide) && !state.declaredDomino) {
		handlePlayerAction(botPlayer.id, 'announce_domino');
		await new Promise(r => setTimeout(r, 1500));
	}
	
	if (canPlayAny(state.hand, game.doubleChainSide)) {
		await new Promise(r => setTimeout(r, 1000));
		
		const validPlays = [];
		state.hand.forEach((t, i) => {
			if (game.leftEnd === null) { validPlays.push({ index: i, side: 'left' }); }
			else {
				if ((!game.doubleChainSide || game.doubleChainSide === 'left') && (t.left === game.leftEnd || t.right === game.leftEnd)) validPlays.push({ index: i, side: 'left' });
				if ((!game.doubleChainSide || game.doubleChainSide === 'right') && (t.left === game.rightEnd || t.right === game.rightEnd)) validPlays.push({ index: i, side: 'right' });
			}
		});
		const choice = validPlays[Math.floor(Math.random() * validPlays.length)];
		handlePlayerAction(botPlayer.id, 'play', choice);
	} else {
		if (game.drawCountThisTurn === 3 || game.deck.length === 0) {
			broadcastAnnounce(`${botPlayer.name} จั่วครบ 3 ตัวแล้ว แต่ยังไม่มีตัวที่วางได้ จบตา`);
			state.declaredDomino = false;
			game.consecutivePasses++;
			if (game.consecutivePasses >= players.length) { handleStuckGame(); return; }
			setTimeout(() => advanceTurn(), 2000);
		}
	}
}

function processTurnLogic() {
	if (!isHost || game.status !== 'playing') return;
	const currentPlayer = players[game.turnIndex];
	const state = game.playerStates[currentPlayer.id];
	
	clearTimeout(botTimer);

	if (game.deck.length === 0 && !canPlayAny(state.hand, game.doubleChainSide)) {
		broadcastAnnounce('กองจั่วหมดและไม่มีผู้เล่นสามารถดำเนินเกมต่อได้', true);
		let minScore = Infinity; let winnerId = null;
		players.forEach(p => {
			let s = 0; game.playerStates[p.id].hand.forEach(t => s += getTileScore(t));
			if(s < minScore) { minScore = s; winnerId = p.id; }
		});
		setTimeout(() => handleWin(winnerId), 2000);
		return;
	}
	
	if (currentPlayer.isBot) {
		runBotTurn(currentPlayer, state);
	} else {
		botTimer = setTimeout(() => {
			if (game.status === 'playing' && players[game.turnIndex].id === currentPlayer.id) {
				const attemptAutoPlay = () => {
					if (game.status !== 'playing' || players[game.turnIndex].id !== currentPlayer.id) return;
					const st = game.playerStates[currentPlayer.id];
					if (canPlayAny(st.hand, game.doubleChainSide)) {
						if (st.hand.length === 2 && !st.declaredDomino) {
							handlePlayerAction(currentPlayer.id, 'announce_domino');
						}
						setTimeout(() => {
							if (game.status !== 'playing' || players[game.turnIndex].id !== currentPlayer.id) return;
							for (let i = 0; i < st.hand.length; i++) {
								const t = st.hand[i];
								let side = null;
								if (game.leftEnd === null) { side = 'left'; }
								else if ((!game.doubleChainSide || game.doubleChainSide === 'left') && (t.left === game.leftEnd || t.right === game.leftEnd)) { side = 'left'; }
								else if ((!game.doubleChainSide || game.doubleChainSide === 'right') && (t.left === game.rightEnd || t.right === game.rightEnd)) { side = 'right'; }
								
								if (side) {
									handlePlayerAction(currentPlayer.id, 'play', { index: i, side: side });
									return;
								}
							}
						}, 500);
					} else if (game.deck.length > 0 && game.drawCountThisTurn < 3) {
						handlePlayerAction(currentPlayer.id, 'draw');
						setTimeout(attemptAutoPlay, 1500);
					} else {
						if (game.drawCountThisTurn < 3 && game.deck.length === 0 && !canPlayAny(st.hand, game.doubleChainSide)) {
							broadcastAnnounce(`กองจั่วหมดแล้ว ${currentPlayer.name} ผ่านรอบ`);
							st.declaredDomino = false;
							game.consecutivePasses++;
							if (game.consecutivePasses >= players.length) { handleStuckGame(); return; }
							setTimeout(() => advanceTurn(), 2000);
						}
					}
				};
				attemptAutoPlay();
			}
		}, 60000);

		// Auto-pass human if deck is empty and cannot play anything
		if (!canPlayAny(state.hand, game.doubleChainSide) && game.deck.length === 0) {
			setTimeout(() => {
				broadcastAnnounce(`กองจั่วหมดแล้ว ${currentPlayer.name} ผ่านรอบ`);
				state.declaredDomino = false;
				game.consecutivePasses++;
				if (game.consecutivePasses >= players.length) { handleStuckGame(); return; }
				setTimeout(() => advanceTurn(), 2000);
			}, 2500);
		}
	}
}

function handleStuckGame() {
	broadcastAnnounce('เกมตัน! ทำการรวมแต้มเพื่อหาผู้ชนะ', true);
	let minScore = Infinity; let winnerId = null;
	players.forEach(p => {
		let s = 0; game.playerStates[p.id].hand.forEach(t => s += getTileScore(t));
		if(s < minScore) { minScore = s; winnerId = p.id; }
	});
	setTimeout(() => handleWin(winnerId), 2000);
}

function handleWin(winnerId) {
	game.status = 'ended';
	const winner = players.find(p => p.id === winnerId);
	
	let resultStats = [];
	players.forEach(p => {
		let pts = 0;
		game.playerStates[p.id].hand.forEach(t => pts += getTileScore(t));
		resultStats.push({ id: p.id, name: p.name, points: pts, isWinner: p.id === winnerId });
	});

	broadcastGameState();
	
	let winnerStat = resultStats.find(r => r.id === winnerId);
	let announceMsg = `การแข่งขันจบแล้ว ${winner.name}เป็นผู้ชนะ เหลือ ${winnerStat.points} แต้ม. `;
	let losers = resultStats.filter(r => r.id !== winnerId);
	losers.forEach(l => {
		announceMsg += `${l.name}ได้ ${l.points} แต้ม แพ้. `;
	});

	broadcastAnnounce(announceMsg, true);
	broadcastSound('win');
	fxCelebrate();

	connections.forEach(c => {
		if(c.open) {
			c.send({ type: 'endGame', winnerName: winner.name, resultStats: resultStats });
			c.send({ type: 'fxCelebrate' });
		}
	});
	showResult(winner.name, resultStats);
}

function showResult(winnerName, resultStats) {
	switchScreen('screen-result', 'title-result');
	stopBGM();
	
	let winnerStat = resultStats.find(r => r.isWinner) || resultStats[0];
	let losers = resultStats.filter(r => !r.isWinner);
	
	let html = `<div style="color: var(--focus-ring); margin-bottom: 10px;">👑 ${winnerName}เป็นผู้ชนะ เหลือ ${winnerStat.points} แต้ม</div>`;
	losers.forEach(l => {
		html += `<div style="color: #ffcdd2; margin-bottom: 10px;">❌ ${l.name}ได้ ${l.points} แต้ม แพ้</div>`;
	});
	
	document.getElementById('winner-text').innerHTML = html;
	document.getElementById('result-status-bar').innerHTML = '';
}

// --- Keyboard Shortcuts ---
document.addEventListener('keydown', (e) => {
	if (!document.getElementById('screen-game').classList.contains('active')) return;
	if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;

	if (e.altKey && e.code === 'KeyC') {
		e.preventDefault();
		const lAria = document.getElementById('board-left-aria').textContent;
		const rAria = document.getElementById('board-right-aria').textContent;
		announce(`${lAria} ${rAria}`);
	}
	else if (e.altKey && e.code === 'KeyK') {
		e.preventDefault();
		const deckAria = document.getElementById('deck-aria-label').textContent;
		announce(deckAria);
	}
	else if (e.altKey && e.code === 'KeyP') {
		e.preventDefault();
		const btnDraw = document.getElementById('btn-draw');
		if (!btnDraw.disabled) btnDraw.click();
	}
	else if (e.altKey && e.code === 'KeyO') {
		e.preventDefault();
		const btnDomino = document.getElementById('btn-domino');
		if (!btnDomino.disabled) btnDomino.click();
	}
	else if (e.altKey && e.code === 'KeyA') {
		e.preventDefault();
		let msg = "สถานะตัวละครทั้งหมด: ";
		const chars = document.querySelectorAll('#top-status-bar .character-card');
		let parts = [];
		chars.forEach(c => {
			const countText = c.querySelector('.char-cards-count').textContent.replace('ไพ่ ', ''); 
			const nameText = c.querySelector('.char-name').textContent.replace(/[🤖👤]/g, '').trim();
			parts.push(`${nameText}เหลือ ${countText}`);
		});
		announce(msg + parts.join(' '));
	}
});