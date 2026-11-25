import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { getDatabase, ref, update, remove, onValue, get } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-database.js";

// --- Language System (UNO Only) ---
const LANG_STORAGE_KEY = 'uno_lang';
let currentLang = 'th';
let lastRoomsSnapshot = null;

function detectInitialLanguage() {
	try {
		const saved = localStorage.getItem(LANG_STORAGE_KEY);
		if (saved === 'th' || saved === 'en') return saved;
	} catch (e) {}
	const navLang = (navigator.language || navigator.userLanguage || 'th').toLowerCase();
	return navLang.startsWith('th') ? 'th' : 'en';
}

function t(thText, enText) {
	return currentLang === 'en' ? enText : thText;
}

function getColorName(colorId) {
	const map = {
		red: { th: 'สีแดง', en: 'Red' },
		blue: { th: 'สีน้ำเงิน', en: 'Blue' },
		green: { th: 'สีเขียว', en: 'Green' },
		yellow: { th: 'สีเหลือง', en: 'Yellow' },
		purple: { th: 'สีม่วง', en: 'Purple' },
		orange: { th: 'สีส้ม', en: 'Orange' }
	};
	const entry = map[colorId];
	if (!entry) return t('ไม่ทราบสี', 'Unknown');
	return entry[currentLang];
}

function getWildColorName(colorId) {
	const map = {
		red: { th: 'แดง', en: 'Red' },
		blue: { th: 'น้ำเงิน', en: 'Blue' },
		green: { th: 'เขียว', en: 'Green' },
		yellow: { th: 'เหลือง', en: 'Yellow' }
	};
	const entry = map[colorId];
	if (!entry) return '';
	return entry[currentLang];
}

function applyLanguageUI() {
	document.documentElement.lang = currentLang;

	const langBtnText = currentLang === 'th' ? 'EN' : 'TH';
	const langBtnAria = currentLang === 'th' ? 'Switch language to English' : 'เปลี่ยนภาษาเป็นภาษาไทย';

	['btn-lang-main', 'btn-lang-lobby'].forEach(id => {
		const btn = document.getElementById(id);
		if (btn) {
			btn.textContent = langBtnText;
			btn.setAttribute('aria-label', langBtnAria);
		}
	});

	const setText = (id, th, en) => {
		const el = document.getElementById(id);
		if (el) el.textContent = t(th, en);
	};
	const setHTML = (id, th, en) => {
		const el = document.getElementById(id);
		if (el) el.innerHTML = t(th, en);
	};
	const setAria = (id, th, en) => {
		const el = document.getElementById(id);
		if (el) el.setAttribute('aria-label', t(th, en));
	};

	// Main screen
	setAria('btn-home-main', 'กลับสู่ชุมชน', 'Back to community');
	setText('subtitle-main-1', 'การ์ดเกมนี้สำหรับการสันทนาการร่วมกัน', 'This card game is for recreation together.');
	setText('subtitle-main-2', 'เกมนี้ออกแบบให้ผู้ใช้ Screen Reader และผู้เล่นทั่วไป สนุกร่วมกันได้อย่างเท่าเทียม', 'Designed so screen reader users and sighted players can enjoy playing together equally.');
	setHTML('btn-create-room', '<span aria-hidden="true">➕</span> สร้างห้องใหม่', '<span aria-hidden="true">➕</span> Create Room');
	setHTML('btn-show-rules', '<span aria-hidden="true">📖</span> วิธีเล่นและกติกา', '<span aria-hidden="true">📖</span> How to Play & Rules');
	setText('heading-open-rooms', 'ห้องที่เปิดอยู่', 'Open Rooms');

	const emptyMsg = document.getElementById('empty-room-msg');
	if (emptyMsg && lastRoomsSnapshot === null) {
		emptyMsg.textContent = t('กำลังโหลดรายการห้อง...', 'Loading rooms...');
	}

	// Rules screen
	setText('title-rules', 'วิธีเล่นและกติกา UNO', 'UNO How to Play & Rules');
	setAria('rules-content-container', 'เนื้อหากติกา', 'Rules content');
	setHTML(
		'rules-goal',
		'<strong>🎯 เป้าหมาย:</strong> ทิ้งไพ่ให้หมดเป็นคนแรก เมื่อเหลือ 2 ใบและมีการ์ดใบหนึ่งใบใดที่วางหรือทิ้งได้ ต้องกด <strong>UNO</strong> (ลืมกดปรับจั่ว 2 ใบ)',
		'<strong>🎯 Goal:</strong> Be the first player to empty your hand. When you have 2 cards left and at least one playable card, you must press <strong>UNO</strong> (forgetting gives a 2-card penalty).'
	);
	setText('rules-heading-play', '🃏 กติกาการลงไพ่', '🃏 Playing Rules');
	setHTML(
		'rules-play-1',
		'ทิ้งไพ่ตาม <strong>สี</strong> หรือ <strong>เลข/ประเภท</strong> ตรงกับใบบนสุดของกองทิ้ง',
		'Play a card that matches the <strong>color</strong> or <strong>number/type</strong> of the top card on the discard pile.'
	);
	setHTML(
		'rules-play-2',
		'ไม่มีไพ่ลงให้กด <strong>จั่วการ์ด</strong> หากยังลงไม่ได้กด <strong>จบตา</strong>',
		'If you have no playable card, press <strong>Draw Card</strong>. If you still cannot play, press <strong>Pass Turn</strong>.'
	);
	setText('rules-heading-special', '✨ ไพ่พิเศษ', '✨ Special Cards');
	setHTML('rules-special-skip', '<strong>Skip:</strong> ข้ามผู้เล่นคนถัดไป', '<strong>Skip:</strong> Skips the next player.');
	setHTML('rules-special-reverse', '<strong>Reverse:</strong> กลับทิศทางการเล่น', '<strong>Reverse:</strong> Reverses the direction of play.');
	setHTML('rules-special-draw2', '<strong>Draw Two (+2):</strong> คนถัดไปจั่ว 2 ใบและถูกข้ามตา', '<strong>Draw Two (+2):</strong> Next player draws 2 cards and loses their turn.');
	setHTML('rules-special-wild', '<strong>Wild:</strong> เลือกเปลี่ยนสีของเกม', '<strong>Wild:</strong> Choose the active color.');
	setHTML('rules-special-wild4', '<strong>Wild Draw Four (+4):</strong> เปลี่ยนสี + คนถัดไปจั่ว 4 ใบ (ท้าทาย/Challenge ได้)', '<strong>Wild Draw Four (+4):</strong> Choose color + next player draws 4 cards (can be Challenged).');
	setText('rules-heading-shortcuts', '⌨️ แป้นพิมพ์ลัด (Keyboard Shortcuts)', '⌨️ Keyboard Shortcuts');
	setHTML('rules-shortcut-p', '<kbd style="background:#334155; padding:2px 5px; border-radius:4px;">Alt + P</kbd> : จั่วการ์ด', '<kbd style="background:#334155; padding:2px 5px; border-radius:4px;">Alt + P</kbd> : Draw Card');
	setHTML('rules-shortcut-x', '<kbd style="background:#334155; padding:2px 5px; border-radius:4px;">Alt + X</kbd> : จบตา / ข้าม', '<kbd style="background:#334155; padding:2px 5px; border-radius:4px;">Alt + X</kbd> : Pass Turn');
	setHTML('rules-shortcut-u', '<kbd style="background:#334155; padding:2px 5px; border-radius:4px;">Alt + U</kbd> : ประกาศ UNO', '<kbd style="background:#334155; padding:2px 5px; border-radius:4px;">Alt + U</kbd> : Declare UNO');
	setHTML('rules-shortcut-c', '<kbd style="background:#334155; padding:2px 5px; border-radius:4px;">Alt + C</kbd> : ฟังการ์ดใบบนสุดกองทิ้ง', '<kbd style="background:#334155; padding:2px 5px; border-radius:4px;">Alt + C</kbd> : Hear top discard card');
	setHTML('rules-shortcut-t', '<kbd style="background:#334155; padding:2px 5px; border-radius:4px;">Alt + T</kbd> : ฟังเวลาที่เหลือในตาเล่น', '<kbd style="background:#334155; padding:2px 5px; border-radius:4px;">Alt + T</kbd> : Hear remaining turn time');
	setText('btn-close-rules', 'กลับหน้าหลัก', 'Back to Main');

	// Lobby screen
	setText('title-lobby-label', 'ล็อบบี้:', 'Lobby:');
	setText('heading-select-color', 'เลือกสีผู้เล่น', 'Select Player Color');
	setAria('color-selection', 'เลือกสี', 'Select color');
	setText('heading-lobby-players-label', 'ผู้เล่นในห้อง', 'Players in Room');
	setText('heading-bot-manage', 'การจัดการ Bot', 'Bot Management');
	setAria('btn-remove-bot', 'ลดบอท', 'Remove bot');
	setAria('btn-add-bot', 'เพิ่มบอท', 'Add bot');
	setHTML('btn-start-game', '<span aria-hidden="true">🚀</span> เริ่มเกม', '<span aria-hidden="true">🚀</span> Start Game');
	setText('btn-leave-lobby', 'ออกจากห้อง', 'Leave Room');

	// Game screen
	setAria('btn-refresh-audio', 'รีเฟรชเสียง', 'Refresh audio');
	setHTML('btn-refresh-audio', '<span aria-hidden="true">🔊</span> รีเฟรช', '<span aria-hidden="true">🔊</span> Refresh');
	setText('deck-unit-visual', 'ใบ', 'cards');
	setText('challenge-prompt-text', 'มีการใช้ Wild Draw Four คุณต้องการ Challenge ไหม?', 'A Wild Draw Four was played. Do you want to Challenge?');
	setText('btn-challenge-yes', 'ใช่ Challenge!', 'Yes, Challenge!');
	setText('btn-challenge-no', 'ไม่ยอมรับ (+4)', 'Decline (+4)');
	setAria('my-cards-container', 'ไพ่ในมือ', 'Cards in hand');
	setText('btn-draw', 'จั่วการ์ด', 'Draw Card');
	setText('btn-pass', 'จบตา', 'Pass Turn');
	setText('color-picker-title', 'เลือกสีสำหรับ Wild', 'Choose a color for Wild');
	setAria('btn-wild-red', 'เลือกสีแดง', 'Choose Red');
	setAria('btn-wild-blue', 'เลือกสีน้ำเงิน', 'Choose Blue');
	setAria('btn-wild-green', 'เลือกสีเขียว', 'Choose Green');
	setAria('btn-wild-yellow', 'เลือกสีเหลือง', 'Choose Yellow');

	// Result screen
	setText('title-result', 'จบเกม!', 'Game Over!');
	setHTML('btn-play-again', '<span aria-hidden="true">🔄</span> เล่นรอบใหม่', '<span aria-hidden="true">🔄</span> Play Again');
	setAria('btn-home-result', 'กลับสู่ชุมชน', 'Back to community');
	setHTML('btn-home-result', '<span aria-hidden="true">🏠</span> กลับหน้าหลัก', '<span aria-hidden="true">🏠</span> Back to Home');

	// Re-render dynamic views if currently visible
	if (document.getElementById('screen-main').classList.contains('active') && lastRoomsSnapshot !== null) {
		renderRoomList(lastRoomsSnapshot);
	}
	if (document.getElementById('screen-lobby').classList.contains('active')) {
		renderLobby();
	}
}

function toggleLanguage() {
	currentLang = currentLang === 'th' ? 'en' : 'th';
	try {
		localStorage.setItem(LANG_STORAGE_KEY, currentLang);
	} catch (e) {}
	applyLanguageUI();
	announce(t('เปลี่ยนเป็นภาษาไทยแล้ว', 'Switched to English'), true);
}

function initLanguageSystem() {
	currentLang = detectInitialLanguage();
	const btnMain = document.getElementById('btn-lang-main');
	const btnLobby = document.getElementById('btn-lang-lobby');
	if (btnMain) btnMain.addEventListener('click', toggleLanguage);
	if (btnLobby) btnLobby.addEventListener('click', toggleLanguage);
	applyLanguageUI();
}

if (document.readyState === 'loading') {
	document.addEventListener('DOMContentLoaded', initLanguageSystem);
} else {
	initLanguageSystem();
}
// ----------------------------------

// --- Web Audio API System ---
const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
const soundBuffers = {};
const soundNames = ['1', 'select', 'start', 'bgm', 'jua', 'turn', 'uno', 'skip', 'reverse', 'draw2', 'draw4', 'll', 'ww', 'wl', 'win', 'abc'];

let turnAudioEnqueued = 0; // ติดตามจำนวนคิวของเสียง turn ที่กำลังจะเล่น/กำลังเล่นอยู่
let pendingABC = false; // ร้องขอการเล่น abc รอไว้

// --- Global Audio Queue ---
const audioQueue = [];
let isAudioPlaying = false;

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

function processAudioQueue() {
	if (isAudioPlaying || audioQueue.length === 0) return;
	isAudioPlaying = true;

	const task = audioQueue.shift();

	const handleEnd = () => {
		if (task.name === 'turn') {
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
		if (task.onEndedCb) task.onEndedCb();
		isAudioPlaying = false;
		processAudioQueue();
	};

	if (!soundBuffers[task.name]) {
		handleEnd();
		return;
	}

	const source = audioCtx.createBufferSource();
	source.buffer = soundBuffers[task.name];
	source.connect(audioCtx.destination);
	
	source.onended = handleEnd;

	try {
		source.start(0);
	} catch (e) {
		console.warn('Audio start failed:', e);
		handleEnd();
	}
}

function playSound(name, onEndedCb = null) {
	if(audioCtx.state === 'suspended') audioCtx.resume();

	if (name === 'bgm') {
		if(!soundBuffers[name]) {
			if(onEndedCb) onEndedCb();
			return null;
		}
		const source = audioCtx.createBufferSource();
		source.buffer = soundBuffers[name];
		source.connect(audioCtx.destination);
		source.loop = true;
		source.onended = () => { if (onEndedCb) onEndedCb(); };
		source.start(0);
		bgmNode = source;
		return source;
	}

	audioQueue.push({ name, onEndedCb });
	processAudioQueue();
	return null; 
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
		playSound('select');
		playSound(soundName);
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
	const colorName = getColorName(playerObj.colorId);
	if (playerObj.id === myPeerId) return t(`คุณ${colorName}`, `You (${colorName})`);
	if (playerObj.isBot) return t(`บอท${colorName}`, `Bot (${colorName})`);
	return t(`เพื่อน${colorName}`, `Friend (${colorName})`);
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
				conn.lastPong = Date.now();
				conn._connectedAt = Date.now();
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
	lastRoomsSnapshot = snapshot.val();
	if (document.getElementById('screen-main').classList.contains('active')) {
		renderRoomList(lastRoomsSnapshot);
	}
});

function renderRoomList(rooms) {
	const list = document.getElementById('room-list');
	const emptyText = t('ไม่มีห้องที่เปิดอยู่', 'No open rooms available');
	if (!rooms) {
		list.innerHTML = `<li id="empty-room-msg" style="text-align: center; color: var(--text-muted);">${emptyText}</li>`;
		return;
	}
	const activeRooms = Object.entries(rooms).filter(([id, r]) => Date.now() - r.lastActive < 120000 && r.status === 'waiting' && r.currentPlayers < MAX_PLAYERS);
	
	if (activeRooms.length === 0) {
		list.innerHTML = `<li id="empty-room-msg" style="text-align: center; color: var(--text-muted);">${emptyText}</li>`;
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
		const hostCountText = t(`มีผู้เล่น ${room.currentPlayers} / ${MAX_PLAYERS} คน`, `${room.currentPlayers} / ${MAX_PLAYERS} players`);
		const roomLabel = t(`ห้อง: ${id}`, `Room: ${id}`);
		const ariaLabel = t(`เข้าร่วมห้อง ${id} ${hostCountText}`, `Join room ${id}, ${hostCountText}`);
		let li = list.querySelector(`li[data-room-id="${id}"]`);
		if (li) {
			const btn = li.querySelector('button');
			if (btn) {
				btn.innerHTML = `<span style="font-weight:bold; font-size:18px;">${roomLabel}</span> <span>${hostCountText}</span>`;
				btn.setAttribute('aria-label', ariaLabel);
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
			btn.innerHTML = `<span style="font-weight:bold; font-size:18px;">${roomLabel}</span> <span>${hostCountText}</span>`;
			btn.setAttribute('aria-label', ariaLabel);
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
		announce(t('ระบบกำลังเตรียมพร้อม กรุณารอสักครู่', 'System is getting ready, please wait.'));
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
		players = [{ id: myPeerId, name: "Host", isBot: false, colorId: myColorId, connection: null }];
		
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
		announce(t('สร้างห้องไม่สำเร็จ กรุณาลองใหม่', 'Failed to create room, please try again.'));
	}
};

function joinRoom(roomId, hostPeerId) {
	if (!myPeerId) {
		announce(t('ระบบกำลังเตรียมพร้อม กรุณารอสักครู่', 'System is getting ready, please wait.'));
		isJoiningRoom = false;
		return;
	}
	
	isHost = false; currentRoomId = roomId;
	announce(t('กำลังเชื่อมต่อไปยัง Host...', 'Connecting to Host...'));
	hostConnection = peer.connect(hostPeerId, { reliable: true });
	
	hostConnection.on('error', (err) => {
		console.error(err);
		announce(t('การเชื่อมต่อล้มเหลว', 'Connection failed'), true);
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
	announce(t('ออกจากห้องแล้ว', 'Left the room'));
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
		const colorLabel = getColorName(c.id);
		const takenSuffix = isTaken ? t(' (ถูกเลือกแล้ว)', ' (Taken)') : '';
		btn.setAttribute('aria-label', colorLabel + takenSuffix);
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
	conn.lastPong = Date.now();
	conn._connectedAt = Date.now();
	conn.on('data', data => {
		if (data.type === 'pong') {
			conn.lastPong = Date.now();
		}
		else if (data.type === 'joinReq') {
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
	conn.on('error', () => handleClientDisconnect(conn.customPeerId || conn.peer));
}

setInterval(() => {
	if (!isHost) return;
	const now = Date.now();
	connections.forEach(c => {
		if (c.open) {
			try {
				c.send({ type: 'ping' });
			} catch (e) {}
		}
		const lastPong = c.lastPong || c._connectedAt || now;
		if (now - lastPong > 12000) {
			const peerId = c.customPeerId || c.peer;
			if (peerId) {
				handleClientDisconnect(peerId);
			}
		}
	});
}, 3000);

document.getElementById('btn-add-bot').onclick = () => {
	if (players.length < MAX_PLAYERS) {
		const colorId = getAvailableColors()[0];
		const botId = 'bot_' + Date.now();
		const colorName = getColorName(colorId);
		players.push({ id: botId, name: `Bot`, isBot: true, colorId: colorId });
		announce(t(`เพิ่มบอท${colorName} 1 ตัวแล้ว`, `Added 1 ${colorName} bot`), true);
		syncLobby();
		broadcastSound('select');
	}
};
document.getElementById('btn-remove-bot').onclick = () => {
	const botIdx = players.slice().reverse().findIndex(p => p.isBot);
	if (botIdx !== -1) {
		const actualIdx = players.length - 1 - botIdx;
		const removedBot = players[actualIdx];
		const colorName = getColorName(removedBot.colorId);
		players.splice(actualIdx, 1);
		announce(t(`ลดบอท${colorName} 1 ตัวแล้ว`, `Removed 1 ${colorName} bot`), true);
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
	if (data.type === 'ping') {
		if (hostConnection && hostConnection.open) {
			hostConnection.send({ type: 'pong' });
		}
	} else if (data.type === 'lobbySync') {
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
			announce(t('เริ่มรอบใหม่แล้ว!', 'New round started!'), true);
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
		const winnerPlayer = players.find(p => p.id === data.winnerId);
		const localizedWinner = winnerPlayer ? getPronounName(winnerPlayer) : (data.winner || t('ผู้เล่น', 'Player'));
		doWinAnimation(localizedWinner, () => {
			showResult(data.winnerId, data.scores, data.matchOver);
		});
	} else if (data.type === 'playSound') {
		playSoundEvent(data.soundName, data.playSelectFirst);
	} else if (data.type === 'startAnim') {
		doStartAnimation(() => {});
	}
}

function handleHostDisconnect() {
	announce(t('Host หลุดการเชื่อมต่อ', 'Host disconnected'), true);
	leaveLobby();
}

function handleClientDisconnect(peerId) {
	if (!isHost) return;
	const p = players.find(x => x.id === peerId);
	if (p && !p.isBot) {
		p.isBot = true;
		broadcastGameEvent({ type: 'playerDisconnected', playerId: peerId });
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

function doWinAnimation(winner, callback) {
	stopBGM();
	const animDiv = document.createElement('div');
	animDiv.id = 'win-anim-screen';
	animDiv.setAttribute('aria-hidden', 'true');
	animDiv.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:radial-gradient(circle at center, #111827 0%, #000000 100%);z-index:10000;display:flex;justify-content:center;align-items:center;color:#ffffff;font-size:3rem;font-weight:bold;text-align:center;text-shadow: 0 0 20px #f7e018;';
	document.body.appendChild(animDiv);
	
	setTimeout(() => { playSound('win'); }, 500);

	const steps = [
		{ time: 200, text: 'Uno' },
		{ time: 1200, text: t('กำลังทำการสรุปผล', 'Calculating results...') },
		{ time: 2500, text: t('ได้แก่ ' + winner, 'The winner is ' + winner) }
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
	broadcastGameEvent({ type: 'gameStarted' });
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
			broadcastGameEvent({ type: 'reshuffleDeck' });
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
		case 'gameStarted': {
			msg = t('เกมเริ่มแล้ว!', 'Game started!');
			break;
		}
		case 'reshuffleDeck': {
			msg = t('สับกองทิ้งใหม่', 'Reshuffling discard pile');
			break;
		}
		case 'playerDisconnected': {
			const p = players.find(x => x.id === data.playerId);
			if (p) {
				const colorName = getColorName(p.colorId);
				msg = t(`เพื่อน${colorName}หลุดการเชื่อมต่อ เปลี่ยนเป็นบอทแล้ว`, `Friend (${colorName}) disconnected and was replaced by a bot`);
				assertive = true;
			}
			break;
		}
		case 'forgotUno': {
			const p = players.find(x => x.id === data.playerId);
			if (p) {
				msg = t(`${getPronounName(p)} ลืมประกาศ UNO โดนปรับจั่ว 2 ใบ`, `${getPronounName(p)} forgot to call UNO and drew 2 penalty cards`);
				assertive = true;
			}
			break;
		}
		case 'declareUno': {
			const p = players.find(x => x.id === data.playerId);
			if (p) {
				msg = t(`${getPronounName(p)} ประกาศ UNO!`, `${getPronounName(p)} called UNO!`);
				assertive = true;
				playSound('uno');
			}
			triggerUnoCinematicAnimation(data.playerId);
			break;
		}
		case 'challengeSuccess': {
			const p = players.find(x => x.id === data.playerId);
			if (p) msg = t(`Challenge สำเร็จ! ${getPronounName(p)} โดนปรับจั่ว 4 ใบ`, `Challenge succeeded! ${getPronounName(p)} drew 4 cards`);
			playSound('ww');
			break;
		}
		case 'challengeFailed': {
			const p = players.find(x => x.id === data.playerId);
			if (p) msg = t(`Challenge พลาด! ${getPronounName(p)} โดนปรับจั่ว 6 ใบ`, `Challenge failed! ${getPronounName(p)} drew 6 cards`);
			playSound('wl');
			break;
		}
		case 'challengeDeclined': {
			const p = players.find(x => x.id === data.playerId);
			if (p) msg = t(`${getPronounName(p)} ไม่ยอมรับ Challenge โดนจั่ว 4 ใบ`, `${getPronounName(p)} declined the Challenge and drew 4 cards`);
			playSound('ll');
			break;
		}
		case 'drawCard': {
			const p = players.find(x => x.id === data.playerId);
			if (p) {
				msg = t(`${getPronounName(p)} จั่วการ์ด 1 ใบ`, `${getPronounName(p)} drew 1 card`);
				playSound('jua');
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
			if (p) msg = t(`${getPronounName(p)} จบตา`, `${getPronounName(p)} passed the turn`);
			turnAudioEnqueued++;
			playSound('turn');
			break;
		}
		case 'winGame': {
			const p = players.find(x => x.id === data.winnerId);
			if (p) {
				msg = t(`จบเกม! ผู้ชนะคือ ${getPronounName(p)}`, `Game over! The winner is ${getPronounName(p)}`);
				assertive = true;
			}
			break;
		}
		case 'playCard': {
			const p = players.find(x => x.id === data.playerId);
			if (!p) break;
			msg = t(`${getPronounName(p)} วางการ์ด ${getCardARIA(data.card)}`, `${getPronounName(p)} played ${getCardARIA(data.card)}`);
			
			if (data.playerId === myPeerId) {
				animateLocalPlay(data.card, data.selectedColor);
			} else {
				animateRemotePlay(data.playerId, data.card, data.selectedColor);
			}

			if (data.card.color === 'wild' && data.selectedColor) {
				const wildColorName = getWildColorName(data.selectedColor);
				msg += t(` เปลี่ยนเป็นสี${wildColorName}`, `, changing color to ${wildColorName}`);
			}

			const sequence = [];
			sequence.push('select');

			if (data.effect === 'skip') {
				sequence.push('skip');
				const target = players.find(x => x.id === data.targetId);
				if (target) msg += t(`, ส่งผลให้ ${getPronounName(target)} ถูกใช้การ์ด skip ข้ามตา`, `, skipping ${getPronounName(target)}'s turn`);
			} else if (data.effect === 'reverse') {
				sequence.push('reverse');
				msg += t(`, ทำการกลับทิศทางการเล่น`, `, reversing the direction of play`);
			} else if (data.effect === 'draw2') {
				sequence.push('draw2');
				const target = players.find(x => x.id === data.targetId);
				if (target) msg += t(`, ส่งผลให้ ${getPronounName(target)} โดนบังคับจั่ว 2 ใบและถูกข้ามตา`, `, forcing ${getPronounName(target)} to draw 2 cards and lose their turn`);
			} else if (data.effect === 'wild4') {
				sequence.push('draw4');
				const target = players.find(x => x.id === data.targetId);
				if (target) msg += t(`, ส่งผลให้ ${getPronounName(target)} ถูกใช้ Wild Draw Four ต้องเลือกว่าจะ Challenge หรือไม่`, `, forcing ${getPronounName(target)} to decide whether to Challenge`);
			}
			
			if (!data.isWin) {
				sequence.push('turn');
				turnAudioEnqueued++;
			}

			sequence.forEach(s => playSound(s));
			break;
		}
		case 'playerJoined': {
			const p = players.find(x => x.id === data.playerId);
			if (p) msg = t(`${getPronounName(p)} เข้าร่วมห้องสำเร็จ`, `${getPronounName(p)} joined the room`);
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
		const winnerPlayer = players.find(p => p.id === winnerId);
		const winnerName = winnerPlayer ? getPronounName(winnerPlayer) : t('ผู้เล่น', 'Player');
		connections.forEach(c => { if (c.open) c.send({ type: 'endGame', winnerId, winner: winnerName, scores: score, matchOver: game.matchOver }); });
		game.status = 'ended'; 
		doWinAnimation(winnerName, () => {
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
				
				if (game.playerStates[currentPlayer.id].hand.length === 0 && game.status !== 'ended') {
					handleWin(currentPlayer.id);
				}
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
			
			if (game.playerStates[currentPlayer.id].hand.length === 0 && game.status !== 'ended') {
				handleWin(currentPlayer.id);
			}
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
	const colorMap = currentLang === 'en'
		? { red: 'Red ', blue: 'Blue ', green: 'Green ', yellow: 'Yellow ', wild: '' }
		: { red: 'แดง ', blue: 'น้ำเงิน ', green: 'เขียว ', yellow: 'เหลือง ', wild: '' };
	const cName = colorMap[card.color] || '';
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
		const playableText = isPlayable ? t('ลงได้', 'playable') : t('ลงไม่ได้', 'not playable');
		btn.setAttribute('aria-label', `${getCardARIA(card)} ${playableText}`);
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
		headingEl.textContent = t(`ถึงรอบ${getPronounName(turnPlayerObj)}`, `${getPronounName(turnPlayerObj)}'s Turn`);
		
		const cDef = COLORS.find(c => c.id === turnPlayerObj.colorId);
		if (cDef) {
			headingEl.style.backgroundColor = cDef.hex;
			if (turnPlayerObj.colorId === 'yellow') {
				headingEl.style.color = '#0f172a';
				headingEl.style.textShadow = '0 1px 2px rgba(255,255,255,0.5)';
				headingEl.style.boxShadow = '0 0 15px rgba(247, 224, 24, 0.6)';
			} else {
				headingEl.style.color = '#ffffff';
				headingEl.style.textShadow = '0 2px 4px rgba(0,0,0,0.7)';
				headingEl.style.boxShadow = `0 0 15px ${cDef.hex}80`;
			}
		}
	}

	if (gameState.turnIndex !== game.lastTurnIndex) {
		const currentPlayer = players[gameState.turnIndex];
		if (currentPlayer) {
			if (currentPlayer.id === myPeerId && !currentPlayer.isBot) {
				if (turnAudioEnqueued > 0) {
					pendingABC = true;
				} else {
					playSound('abc');
				}
			} else {
				pendingABC = false;
			}

			const pronoun = getPronounName(currentPlayer);
			if (currentPlayer.id === myPeerId) {
				let announceMsg = t(`ถึงรอบ${pronoun}แล้ว`, `It is ${pronoun}'s turn.`);
				if (gameState.topDiscard) {
					const displayCard = {...gameState.topDiscard};
					if (displayCard.color === 'wild' && gameState.currentColor) {
						displayCard.color = gameState.currentColor;
					}
					let cardName = getCardARIA(displayCard);
					announceMsg += t(` ขณะนี้การ์ด ${cardName} กำลังแสดงอยู่`, ` Top card is ${cardName}.`);
				}
				announce(announceMsg);
			} else {
				announce(t(`ถึงรอบ${pronoun}แล้ว`, `It is ${pronoun}'s turn.`));
			}
		}
		game.lastTurnIndex = gameState.turnIndex;
	}

	const myHand = gameState.myHand || [];
	if (typeof window._myPrevHand !== 'undefined') {
		if (myHand.length > window._myPrevHand.length) {
			const diff = myHand.length - window._myPrevHand.length;
			const newCards = myHand.slice(-diff);
			newCards.forEach(c => { announce(t(`ได้รับการ์ด ${getCardARIA(c)}`, `Received ${getCardARIA(c)}`)); });
		}
	}
	window._myPrevHand = [...myHand];

	const statusBar = document.getElementById('top-status-bar');
	statusBar.innerHTML = '';
	
	const dirText = gameState.direction === 1 ? t('ขวา', 'right') : t('ซ้าย', 'left');
	let ariaStatusBarText = t(`ทิศทางการเล่น ${dirText}, `, `Direction ${dirText}, `);
	const dirDiv = document.createElement('div');
	dirDiv.className = 'direction-indicator';
	dirDiv.setAttribute('aria-hidden', 'true');
	dirDiv.style.transform = gameState.direction === 1 ? 'rotate(0deg)' : 'rotate(180deg)';
	dirDiv.textContent = '➡️';
	statusBar.appendChild(dirDiv);

	players.forEach((p, idx) => {
		const pState = gameState.playerStates[p.id];
		const isTurn = idx === gameState.turnIndex;
		const cDef = COLORS.find(c => c.id === p.colorId);
		const pName = getPronounName(p);
		
		if (pState) {
			ariaStatusBarText += t(
				`${pName} มีการ์ด ${pState.cardCount} ใบ คะแนน ${pState.score}. `,
				`${pName} has ${pState.cardCount} cards, score ${pState.score}. `
			);
			const charDiv = document.createElement('div');
			charDiv.className = `character-card ${isTurn ? 'is-turn' : ''}`;
			charDiv.setAttribute('data-player-id', p.id);
			charDiv.setAttribute('aria-hidden', 'true');
			charDiv.innerHTML = `
				<div class="char-cards-count">${pState.cardCount} ${t('ใบ', 'cards')}</div>
				<div class="char-avatar" style="background-color: ${cDef?cDef.hex:'#fff'};"><span aria-hidden="true">${p.isBot?'🤖':getMascotEmoji(p.colorId)}</span></div>
				<div class="char-score">${t('คะแนน:', 'Score:')} ${pState.score}</div>
			`;
			statusBar.appendChild(charDiv);
		}
	});
	statusBar.setAttribute('aria-label', ariaStatusBarText);

	document.getElementById('deck-count-visual').textContent = gameState.deckCount;
	document.getElementById('deck-aria-label').textContent = t(
		`การ์ดในกองเหลือ ${gameState.deckCount} ใบ`,
		`${gameState.deckCount} cards remaining in deck`
	);

	const discardEl = document.getElementById('discard-pile');
	discardEl.innerHTML = '';
	if (gameState.topDiscard) {
		const displayCard = {...gameState.topDiscard};
		if (displayCard.color === 'wild' && gameState.currentColor) displayCard.color = gameState.currentColor;
		discardEl.appendChild(renderCardHTML(displayCard));
		document.getElementById('discard-aria-label').textContent = t(
			`การ์ดที่แสดงคือ ${getCardARIA(displayCard)}`,
			`Top discard card is ${getCardARIA(displayCard)}`
		);
	}

	const chalArea = document.getElementById('challenge-area');
	if (gameState.challengeData && gameState.challengeData.victimId === myPeerId) {
		chalArea.style.display = 'block';
		toggleMainGameUI(false);
		document.getElementById('btn-challenge-yes').onclick = () => sendAction('challenge_yes');
		document.getElementById('btn-challenge-no').onclick = () => sendAction('challenge_no');
		setTimeout(() => document.getElementById('btn-challenge-yes').focus(), 100);
	} else {
		chalArea.style.display = 'none';
		if (document.getElementById('color-picker-modal').style.display !== 'flex') {
			toggleMainGameUI(true);
		}
	}

	const myContainer = document.getElementById('my-cards-container');
	myContainer.innerHTML = '';
	
	if (myHand.length > 8) myContainer.classList.add('compact');
	else myContainer.classList.remove('compact');

	const isMyTurn = turnPlayerObj && turnPlayerObj.id === myPeerId && !gameState.challengeData;
	
	myHand.forEach((card, idx) => {
		const playable = isMyTurn && isValidPlay(card, myHand);
		myContainer.appendChild(renderCardHTML(card, idx, playable));
	});

	const btnDraw = document.getElementById('btn-draw');
	const btnPass = document.getElementById('btn-pass');
	const btnUno = document.getElementById('btn-uno');
	
	const myState = gameState.playerStates[myPeerId];
	const playableCardsCount = myHand.filter(card => isValidPlay(card, myHand)).length;

	if (isMyTurn) {
		if (myState && myState.declaredUNO && myHand.length === 2) {
			btnDraw.disabled = true;
			btnPass.disabled = true;
		} else if (myState && myState.hasDrawn) {
			btnDraw.disabled = true;
			btnPass.disabled = false;
			btnPass.onclick = () => { sendAction('pass'); };
		} else {
			btnDraw.disabled = false;
			btnPass.disabled = true;
			btnDraw.onclick = () => { sendAction('draw'); };
		}
	} else {
		btnDraw.disabled = true;
		btnPass.disabled = true;
	}

	if (isMyTurn && myHand.length === 2 && myState && !myState.declaredUNO && playableCardsCount > 0) {
		btnUno.disabled = false;
		btnUno.onclick = () => {
			btnUno.disabled = true;
			sendAction('announce_uno');
		};
	} else {
		btnUno.disabled = true;
	}
}

let pendingPlayCardIndex = -1;
window.selectWildColor = function(color) {
	document.getElementById('color-picker-modal').style.display = 'none';
	toggleMainGameUI(true);
	if (pendingPlayCardIndex !== -1) {
		sendAction('play', { cardIndex: pendingPlayCardIndex, selectedColor: color });
		pendingPlayCardIndex = -1;
	}
};

function onCardClicked(index, card) {
	if (card.color === 'wild') {
		playSound('select'); 
		pendingPlayCardIndex = index;
		const modal = document.getElementById('color-picker-modal');
		modal.style.display = 'flex';
		toggleMainGameUI(false);
		setTimeout(() => modal.querySelector('.color-btn').focus(), 100);
	} else {
		sendAction('play', { cardIndex: index });
	}
}

function sendAction(action, payload = null) {
	if (isHost) handlePlayerAction(myPeerId, action, payload);
	else if (hostConnection && hostConnection.open) hostConnection.send({ type: 'action', action, payload });
}

function showResult(winnerId, score, matchOver = false) {
	switchScreen('screen-result', 'title-result');
	const winner = players.find(p => p.id === winnerId);
	
	let winnerHtml = t(
		`<span aria-hidden="true">🏆</span> ผู้ชนะ: ${getPronounName(winner)} (ได้คะแนนเพิ่ม: ${score})`,
		`<span aria-hidden="true">🏆</span> Winner: ${getPronounName(winner)} (+${score} pts)`
	);
	
	if (matchOver && winner) {
		const colorName = getColorName(winner.colorId);
		const kingTexts = currentLang === 'en' ? [
			`${colorName} is the King of UNO!`,
			`Congratulations! ${colorName} has become the King of UNO!`,
			`Match complete! ${colorName} claims the UNO crown!`,
			`Unstoppable! ${colorName} is the true UNO Champion!`,
			`Hail the champion! ${colorName} rules this UNO room!`
		] : [
			`ตัวละคร${colorName} คือจ้าวแห่ง Uno!`,
			`ขอแสดงความยินดี! ตัวละคร${colorName} ได้ก้าวขึ้นเป็นจ้าวแห่ง Uno!`,
			`เกมจบสมบูรณ์! ตัวละคร${colorName} คว้าตำแหน่งจ้าวแห่ง Uno ไปครอง!`,
			`ไม่มีใครหยุดได้! ตัวละคร${colorName} คือจ้าวแห่ง Uno ตัวจริง!`,
			`ครองบัลลังก์สำเร็จ ตัวละคร${colorName} คือจ้าวแห่ง Uno แห่งห้องนี้!`
		];
		const randText = kingTexts[Math.floor(Math.random() * kingTexts.length)];
		winnerHtml += `<br><br><span style="color: var(--focus-ring); font-size: 32px; font-weight: bold; text-shadow: 0 2px 10px rgba(255,215,0,0.5);">${randText}</span>`;
	}

	let scoreListHtml = `<div style="margin-top: 25px; font-size: 20px; color: var(--text-main);">
		<div style="font-weight: bold; margin-bottom: 10px;">${t('คะแนนสะสมปัจจุบันของผู้เล่นทุกคน', 'Current Total Scores')}</div>
		<ul style="list-style: none; padding: 0;">`;
	players.forEach(p => {
		const colorName = getColorName(p.colorId);
		const pScore = game.playerStates[p.id] ? game.playerStates[p.id].score : 0;
		scoreListHtml += `<li style="margin-bottom: 5px;">${t(`${colorName}: คะแนนสะสม ${pScore} แต้ม`, `${colorName}: ${pScore} pts`)}</li>`;
	});
	scoreListHtml += `</ul></div>`;
	
	winnerHtml += scoreListHtml;

	document.getElementById('winner-text').innerHTML = winnerHtml;
	
	const resultStatusBar = document.getElementById('result-status-bar');
	if (resultStatusBar) {
		resultStatusBar.innerHTML = '';
		players.forEach(p => {
			const pState = game.playerStates[p.id];
			const cDef = COLORS.find(c => c.id === p.colorId);
			if (pState) {
				const charDiv = document.createElement('div');
				charDiv.className = `character-card`;
				charDiv.setAttribute('data-player-id', p.id);
				charDiv.setAttribute('aria-hidden', 'true');
				const cardCount = pState.cardCount !== undefined ? pState.cardCount : (pState.hand ? pState.hand.length : 0);
				charDiv.innerHTML = `
					<div class="char-cards-count">${cardCount} ${t('ใบ', 'cards')}</div>
					<div class="char-avatar" style="background-color: ${cDef?cDef.hex:'#fff'};"><span aria-hidden="true">${p.isBot?'🤖':getMascotEmoji(p.colorId)}</span></div>
					<div class="char-score">${t('คะแนน:', 'Score:')} ${pState.score}</div>
				`;
				resultStatusBar.appendChild(charDiv);
			}
		});
	}

	if (isHost) {
		document.getElementById('btn-start-game').style.display = 'block'; 
		if (matchOver) {
			document.getElementById('result-host-controls').style.display = 'none';
		} else {
			document.getElementById('result-host-controls').style.display = 'flex';
			document.getElementById('btn-play-again').onclick = () => {
				connections.forEach(c => { if(c.open) c.send({ type: 'startAnim' }); });
				doStartAnimation(() => {
					initUNOGame();
					setTimeout(() => { playSound('bgm'); connections.forEach(c => { if(c.open) c.send({ type: 'playSound', soundName: 'bgm' }); }); }, 200);
				});
			};
		}
	}
}

function showAnimOverlay(text) {
	if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
	const overlay = document.getElementById('anim-overlay');
	const textEl = document.getElementById('anim-text');
	textEl.textContent = text;
	overlay.style.display = 'flex';
	setTimeout(() => { overlay.style.display = 'none'; }, 1500);
}

document.getElementById('btn-refresh-audio').onclick = () => {
	 announce(t('รีเฟรชระบบเสียงแล้ว', 'Audio system refreshed'));
	 if(audioCtx.state === 'suspended') audioCtx.resume();
};

document.addEventListener('keydown', (e) => {
	if (e.altKey && e.key === 'p') {
		const btnDraw = document.getElementById('btn-draw');
		if(btnDraw && !btnDraw.disabled && game.status === 'playing') btnDraw.click();
	}
	if (e.altKey && e.key === 'x') {
		const btnPass = document.getElementById('btn-pass');
		if(btnPass && !btnPass.disabled && game.status === 'playing') btnPass.click();
	}
	if (e.altKey && e.key === 'u') {
		const btnUno = document.getElementById('btn-uno');
		if(btnUno && !btnUno.disabled && game.status === 'playing') btnUno.click();
	}
	if (e.altKey && e.key.toLowerCase() === 'c') {
		e.preventDefault();
		if(game.status === 'playing') {
			const ariaLabel = document.getElementById('discard-aria-label');
			if (ariaLabel && ariaLabel.textContent) {
				announce(ariaLabel.textContent);
			}
		}
	}
	if (e.altKey && e.key === 't') {
		if(game.status === 'playing') {
			const elapsed = Math.floor((Date.now() - game.currentTurnStartTime) / 1000);
			let rem = 40 - elapsed;
			if(rem < 0) rem = 0;
			announce(t(`เหลือเวลา ${rem} วินาที`, `${rem} seconds remaining`));
		}
	}
});