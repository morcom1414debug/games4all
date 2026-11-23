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

// ==================== Bilingual Support (TH / EN) ====================
const LANG_STORAGE_KEY = 'crazy_lang';

function detectInitialLanguage() {
	try {
		const saved = localStorage.getItem(LANG_STORAGE_KEY);
		if (saved === 'th' || saved === 'en') return saved;
	} catch (e) {}
	const browserLang = (navigator.language || navigator.userLanguage || 'en').toLowerCase();
	return browserLang.startsWith('th') ? 'th' : 'en';
}

let currentLang = detectInitialLanguage();

const CRAZY_I18N = {
	th: {
		back_community_aria: "กลับสู่ชุมชน",
		login_subtitle: "สำหรับการสันทนาการร่วมกัน",
		login_a11y_desc: "เกมนี้ออกแบบให้ผู้ใช้ Screen Reader และผู้เล่นทั่วไป สนุกร่วมกันได้อย่างเท่าเทียม",
		label_nickname: "กรอกชื่อเล่นของคุณ",
		placeholder_nickname: "กรอกชื่อเล่นของคุณ",
		btn_confirm_name: "ยืนยันชื่อ",
		hello_prefix: "สวัสดี",
		btn_create_room: "สร้างห้องใหม่",
		btn_show_rules: "คู่มือการเล่น",
		open_rooms_heading: "ห้องที่เปิดอยู่",
		loading_rooms: "กำลังโหลดรายการห้อง...",
		rules_title: "คู่มือการเล่น Crazy Eights",
		rules_content_aria: "เนื้อหาคู่มือ",
		rules_html: `
            <p style="font-size: 16px; margin-bottom: 20px; background: rgba(255,255,255,0.05); padding: 12px; border-radius: 8px;"><strong>🎯 เป้าหมาย:</strong> ทิ้งการ์ดให้หมดมือเป็นคนแรก เพื่อเป็นผู้ชนะ</p>
            
            <h3 style="font-size: 16px; margin-top: 15px; text-align: left; border-left: 4px solid var(--primary); padding-left: 12px;">กติกาการวางการ์ด</h3>
            <ul style="margin-left: 20px; margin-bottom: 25px; color: var(--text-muted); line-height: 1.8;">
                <li>ผู้เล่นลงการ์ดที่มี <strong style="color:#fff;">ดอก (Suit)</strong> หรือ <strong style="color:#fff;">ค่า (Rank)</strong> ตรงกับการ์ดบนสุดของกองทิ้ง</li>
                <li>หากไม่มีการ์ดลง ต้อง <strong style="color:var(--info);">กดปุ่มจั่ว</strong> (จั่ว 1 ใบต่อเทิร์น แล้วข้ามตาอัตโนมัติ)</li>
                <li>การ์ด <strong style="color:var(--focus-ring);">เลข 8</strong> สามารถลงทับการ์ดใดก็ได้ทุกกรณี</li>
                <li>เมื่อลงเลข 8 ผู้เล่นสามารถประกาศ <strong style="color:#fff;">เปลี่ยนดอกการ์ด</strong> เป็นดอกใดก็ได้</li>
            </ul>
            
            <h3 style="font-size: 16px; margin-top: 15px; text-align: left; border-left: 4px solid var(--info); padding-left: 12px;">คีย์ลัด (Keyboard Shortcuts)</h3>
            <ul style="margin-left: 20px; margin-bottom: 15px; color: var(--text-muted); line-height: 1.8;">
                <li><strong style="color:#fff; background: rgba(0,0,0,0.5); padding: 2px 6px; border-radius: 4px;">Alt+P</strong> = จั่ว</li>
                <li><strong style="color:#fff; background: rgba(0,0,0,0.5); padding: 2px 6px; border-radius: 4px;">Alt+C</strong> = ฟังข้อมูลการ์ดบนกองทิ้ง</li>
                <li><strong style="color:#fff; background: rgba(0,0,0,0.5); padding: 2px 6px; border-radius: 4px;">Alt+K</strong> = ฟังจำนวนการ์ดที่เหลือในกองจั่ว</li>
                <li><strong style="color:#fff; background: rgba(0,0,0,0.5); padding: 2px 6px; border-radius: 4px;">Alt+A</strong> = ฟังจำนวนการ์ดที่เหลือของผู้เล่นแต่ละคน</li>
            </ul>
		`,
		btn_back_main: "กลับหน้าหลัก",
		lobby_prefix: "ล็อบบี้:",
		lobby_ready_msg: "ห้องพร้อมแล้ว รอเพื่อนหรือเพิ่มบอทได้ทันที",
		players_in_room: "ผู้เล่นในห้อง",
		bot_management_heading: "การจัดการ Bot",
		btn_remove_bot: "ลดบอท",
		btn_remove_bot_aria: "ลดบอท",
		btn_add_bot: "เพิ่มบอท",
		btn_add_bot_aria: "เพิ่มบอท",
		btn_start_game: "🚀 เริ่มเกม",
		btn_leave_lobby: "ออกจากห้อง",
		btn_refresh_audio_text: "รีเฟรช",
		btn_refresh_audio_aria: "รีเฟรชเสียง",
		my_cards_aria: "การ์ดในมือ",
		btn_draw: "จั่ว",
		suit_picker_title: "เลือกดอกถัดไป",
		suit_hearts_aria: "เปลี่ยนเป็นดอก โพแดง",
		suit_diamonds_aria: "เปลี่ยนเป็นดอก ข้าวหลามตัด",
		suit_clubs_aria: "เปลี่ยนเป็นดอก ดอกจิก",
		suit_spades_aria: "เปลี่ยนเป็นดอก โพดำ",
		result_title: "ผลการแข่งขัน"
	},
	en: {
		back_community_aria: "Back to Community",
		login_subtitle: "For Shared Recreation & Fun",
		login_a11y_desc: "Designed for Screen Reader users and sighted players to enjoy equally together.",
		label_nickname: "Enter your nickname",
		placeholder_nickname: "Enter your nickname",
		btn_confirm_name: "Confirm Name",
		hello_prefix: "Hello",
		btn_create_room: "Create New Room",
		btn_show_rules: "How to Play",
		open_rooms_heading: "Open Rooms",
		loading_rooms: "Loading rooms...",
		rules_title: "Crazy Eights Game Manual",
		rules_content_aria: "Game manual content",
		rules_html: `
            <p style="font-size: 16px; margin-bottom: 20px; background: rgba(255,255,255,0.05); padding: 12px; border-radius: 8px;"><strong>🎯 Objective:</strong> Be the first player to discard all cards in your hand to win.</p>
            
            <h3 style="font-size: 16px; margin-top: 15px; text-align: left; border-left: 4px solid var(--primary); padding-left: 12px;">Card Placement Rules</h3>
            <ul style="margin-left: 20px; margin-bottom: 25px; color: var(--text-muted); line-height: 1.8;">
                <li>Play a card matching either the <strong style="color:#fff;">Suit</strong> or <strong style="color:#fff;">Rank</strong> of the top card on the discard pile.</li>
                <li>If you have no playable cards, press <strong style="color:var(--info);">Draw</strong> (draw 1 card per turn; if unplayable, turn passes automatically).</li>
                <li>An <strong style="color:var(--focus-ring);">8 card</strong> can be played on top of any card at any time.</li>
                <li>When playing an 8, you may choose the <strong style="color:#fff;">next active suit</strong>.</li>
            </ul>
            
            <h3 style="font-size: 16px; margin-top: 15px; text-align: left; border-left: 4px solid var(--info); padding-left: 12px;">Keyboard Shortcuts</h3>
            <ul style="margin-left: 20px; margin-bottom: 15px; color: var(--text-muted); line-height: 1.8;">
                <li><strong style="color:#fff; background: rgba(0,0,0,0.5); padding: 2px 6px; border-radius: 4px;">Alt+P</strong> = Draw a card</li>
                <li><strong style="color:#fff; background: rgba(0,0,0,0.5); padding: 2px 6px; border-radius: 4px;">Alt+C</strong> = Announce top card on the discard pile</li>
                <li><strong style="color:#fff; background: rgba(0,0,0,0.5); padding: 2px 6px; border-radius: 4px;">Alt+K</strong> = Announce remaining cards in the draw pile</li>
                <li><strong style="color:#fff; background: rgba(0,0,0,0.5); padding: 2px 6px; border-radius: 4px;">Alt+A</strong> = Announce card counts for all players</li>
            </ul>
		`,
		btn_back_main: "Back to Home",
		lobby_prefix: "Lobby:",
		lobby_ready_msg: "Room is ready! Wait for friends or add bots right away.",
		players_in_room: "Players in Room",
		bot_management_heading: "Bot Management",
		btn_remove_bot: "Remove Bot",
		btn_remove_bot_aria: "Remove bot",
		btn_add_bot: "Add Bot",
		btn_add_bot_aria: "Add bot",
		btn_start_game: "🚀 Start Game",
		btn_leave_lobby: "Leave Room",
		btn_refresh_audio_text: "Refresh",
		btn_refresh_audio_aria: "Refresh audio",
		my_cards_aria: "Cards in hand",
		btn_draw: "Draw",
		suit_picker_title: "Choose Next Suit",
		suit_hearts_aria: "Change suit to Hearts",
		suit_diamonds_aria: "Change suit to Diamonds",
		suit_clubs_aria: "Change suit to Clubs",
		suit_spades_aria: "Change suit to Spades",
		result_title: "Match Results"
	}
};

const BOT_NAME_EN_MAP = {
	'บอทสายฟ้า': 'Bot Lightning',
	'บอทเจ้าป่า': 'Bot JungleKing',
	'บอทดาวเหนือ': 'Bot NorthStar',
	'บอทขุนพล': 'Bot General',
	'บอทจอมทัพ': 'Bot Warlord',
	'บอทพายุ': 'Bot Storm',
	'บอทฟีนิกซ์': 'Bot Phoenix',
	'บอทนักรบ': 'Bot Warrior',
	'บอทเสือดำ': 'Bot BlackPanther',
	'บอทราชัน': 'Bot Monarch',
	'บอทมังกร': 'Bot Dragon',
	'บอทภูผา': 'Bot Mountain',
	'บอททะเล': 'Bot Ocean',
	'บอทวายุ': 'Bot Zephyr',
	'บอทหมอก': 'Bot Mist',
	'บอทตะวัน': 'Bot Sun',
	'บอทจันทรา': 'Bot Moon',
	'บอทแสงดาว': 'Bot Starlight',
	'บอทเพชร': 'Bot Diamond',
	'บอทโชคดี': 'Bot Lucky'
};

function formatPlayerDisplayName(rawName) {
	if (!rawName) return '';
	if (currentLang !== 'en') return rawName;
	if (BOT_NAME_EN_MAP[rawName]) return BOT_NAME_EN_MAP[rawName];
	if (rawName.startsWith('บอท_')) return 'Bot_' + rawName.slice(4);
	if (rawName.startsWith('บอท')) return 'Bot ' + rawName.slice(3);
	return rawName;
}

function replaceAllBotNamesToEn(text) {
	let s = String(text);
	for (const [thName, enName] of Object.entries(BOT_NAME_EN_MAP)) {
		s = s.split(thName).join(enName);
	}
	s = s.replace(/บอท_/g, 'Bot_').replace(/บอท([^\s]+)/g, 'Bot $1');
	return s;
}

function formatCardNameForLang(card) {
	if (!card) return '';
	if (currentLang !== 'en') return card.name;
	const suitEnMap = { '♥': 'Hearts', '♦': 'Diamonds', '♣': 'Clubs', '♠': 'Spades' };
	const rankEnMap = { 'A': 'Ace', 'J': 'Jack', 'Q': 'Queen', 'K': 'King' };
	const rText = rankEnMap[card.rank] || card.rank;
	const sText = suitEnMap[card.suit] || card.suit;
	return `${rText} of ${sText}`;
}

function translateCardPhraseToEn(phrase) {
	let s = String(phrase);
	s = s.replace(/เอซ/g, 'Ace')
		.replace(/แจ็ค/g, 'Jack')
		.replace(/แหม่ม/g, 'Queen')
		.replace(/คิง/g, 'King');
	s = s.replace(/\b(Ace|2|3|4|5|6|7|8|9|10|Jack|Queen|King)\s+(โพแดง|ข้าวหลามตัด|ดอกจิก|โพดำ)/g, (_, r, suitTh) => {
		const suitEn = { 'โพแดง': 'Hearts', 'ข้าวหลามตัด': 'Diamonds', 'ดอกจิก': 'Clubs', 'โพดำ': 'Spades' }[suitTh] || suitTh;
		return `${r} of ${suitEn}`;
	});
	s = s.replace(/โพแดง/g, 'Hearts')
		.replace(/ข้าวหลามตัด/g, 'Diamonds')
		.replace(/ดอกจิก/g, 'Clubs')
		.replace(/โพดำ/g, 'Spades');
	return s;
}

function translateAnnouncementToEn(rawMsg) {
	if (!rawMsg) return '';
	let s = replaceAllBotNamesToEn(rawMsg);

	// 1. ข้อความเริ่มรอบ (broadcastTurnStart)
	const turnMatch = s.match(/^ถึงรอบของ (.+?) การ์ดกองทิ้งคือ (.+)$/);
	if (turnMatch) {
		const pName = turnMatch[1];
		let cardDesc = turnMatch[2];
		cardDesc = cardDesc.replace(/ เปลี่ยนดอกเป็น (โพแดง|ข้าวหลามตัด|ดอกจิก|โพดำ)$/, (_, suitTh) => {
			return ` (Suit changed to ${translateCardPhraseToEn(suitTh)})`;
		});
		cardDesc = translateCardPhraseToEn(cardDesc);
		return `${pName}'s turn. Discard pile: ${cardDesc}`;
	}

	// 2. ข้อความล็อบบี้และการเชื่อมต่อ
	s = s.replace(/^เพื่อน(.+?) หลุดการเชื่อมต่อ เปลี่ยนเป็นบอทแล้ว$/, '$1 disconnected and was replaced by a bot.');
	s = s.replace(/^(.+?) เข้าร่วมห้องสำเร็จ$/, '$1 joined the room.');
	s = s.replace(/^เพิ่มบอท (.+?) เข้าห้องแล้ว$/, 'Added $1 to the room.');
	s = s.replace(/^ลด (.+?) ออกจากห้องแล้ว$/, 'Removed $1 from the room.');

	// 3. ข้อความการจั่ว
	const privDrawMatch = s.match(/^คุณจั่วได้การ์ด (.+)$/);
	if (privDrawMatch) {
		return `You drew ${translateCardPhraseToEn(privDrawMatch[1])}`;
	}
	s = s.replace(/^(.+?) จั่วการ์ด 1 ใบ$/, '$1 drew 1 card.');
	if (s === 'ไม่มีการ์ดลงได้ จบตา') {
		return 'No playable cards. Turn ended.';
	}

	// 4. ข้อความลงการ์ด
	const play8Match = s.match(/^(.+?) ลง (.+?) และเปลี่ยนดอกเป็น (โพแดง|ข้าวหลามตัด|ดอกจิก|โพดำ)$/);
	if (play8Match) {
		return `${play8Match[1]} played ${translateCardPhraseToEn(play8Match[2])} and changed the suit to ${translateCardPhraseToEn(play8Match[3])}.`;
	}
	const playMatch = s.match(/^(.+?) ลง (.+)$/);
	if (playMatch) {
		return `${playMatch[1]} played ${translateCardPhraseToEn(playMatch[2])}.`;
	}

	// 5. เหลือ 1 ใบ / สับกองใหม่ / จบเกม
	s = s.replace(/^(.+?) เหลือ 1 ใบ$/, '$1 has 1 card left!');
	if (s === 'กองจั่วหมด สับการ์ดใหม่เรียบร้อยแล้ว') {
		return 'Draw pile empty. Shuffled discard pile into a new deck.';
	}
	const winMatch = s.match(/^การแข่งขันจบแล้ว (.+?) เป็นผู้ชนะ\s*$/);
	if (winMatch) {
		return `Match over! ${winMatch[1]} is the winner!`;
	}

	return s;
}

function getLocalizedMessage(msg) {
	if (!msg) return '';
	return currentLang === 'en' ? translateAnnouncementToEn(msg) : msg;
}

let lastRenderedRooms = null;
let lastRenderedGameState = null;
let lastDisplayedWinner = '';
let lastDisplayedStats = null;
let suitModalCancelBtn = null;

function applyStaticLanguage() {
	const dict = CRAZY_I18N[currentLang] || CRAZY_I18N.th;
	document.documentElement.lang = currentLang;

	document.querySelectorAll('[data-i18n]').forEach(el => {
		const k = el.getAttribute('data-i18n');
		if (dict[k] !== undefined) el.textContent = dict[k];
	});

	document.querySelectorAll('[data-i18n-html]').forEach(el => {
		const k = el.getAttribute('data-i18n-html');
		if (dict[k] !== undefined) el.innerHTML = dict[k];
	});

	document.querySelectorAll('[data-i18n-placeholder]').forEach(el => {
		const k = el.getAttribute('data-i18n-placeholder');
		if (dict[k] !== undefined) el.setAttribute('placeholder', dict[k]);
	});

	document.querySelectorAll('[data-i18n-aria]').forEach(el => {
		const k = el.getAttribute('data-i18n-aria');
		if (dict[k] !== undefined) el.setAttribute('aria-label', dict[k]);
	});

	if (suitModalCancelBtn) {
		suitModalCancelBtn.textContent = currentLang === 'en' ? 'Cancel' : 'ยกเลิก';
		suitModalCancelBtn.setAttribute('aria-label', currentLang === 'en' ? 'Cancel playing 8 card' : 'ยกเลิกการลงการ์ด 8');
	}

	const label = currentLang === 'th' ? '🌐 EN' : '🌐 ไทย';
	const aria = currentLang === 'th' ? 'Switch language to English' : 'สลับภาษาเป็นภาษาไทย';
	['btn-lang-login', 'btn-lang-main'].forEach(id => {
		const b = document.getElementById(id);
		if (b) {
			b.textContent = label;
			b.setAttribute('aria-label', aria);
		}
	});
}

function toggleLanguage() {
	currentLang = currentLang === 'th' ? 'en' : 'th';
	try {
		localStorage.setItem(LANG_STORAGE_KEY, currentLang);
	} catch (e) {}
	applyStaticLanguage();

	if (lastRenderedRooms !== null && document.getElementById('screen-main').classList.contains('active')) {
		renderRoomList(lastRenderedRooms);
	}
	if (document.getElementById('screen-lobby').classList.contains('active')) {
		renderLobby();
	}
	if (lastRenderedGameState && document.getElementById('screen-game').classList.contains('active')) {
		renderGame(lastRenderedGameState);
	}
	if (lastDisplayedStats && document.getElementById('screen-result').classList.contains('active')) {
		showResult(lastDisplayedWinner, lastDisplayedStats);
	}
	announce(currentLang === 'en' ? 'Language switched to English' : 'เปลี่ยนภาษาเป็นภาษาไทยแล้ว');
}

// --- Audio System (Unchanged Timing & Logic) ---
const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
const soundBuffers = {};
const soundNames = ['1', 'select', 'start', 'bgm', 'jua', 'turn', 'uno', 'win', 'hit', 'abc'];
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
	const localMsg = getLocalizedMessage(msg);
	announce(localMsg, assertive);
	logEvent(localMsg);
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
let previousTurnIndex = -1;

const MAX_PLAYERS = 4;
const BOT_NAMES = [
	'บอทสายฟ้า', 'บอทเจ้าป่า', 'บอทดาวเหนือ', 'บอทขุนพล', 'บอทจอมทัพ', 
	'บอทพายุ', 'บอทฟีนิกซ์', 'บอทนักรบ', 'บอทเสือดำ', 'บอทราชัน',
	'บอทมังกร', 'บอทภูผา', 'บอททะเล', 'บอทวายุ', 'บอทหมอก', 
	'บอทตะวัน', 'บอทจันทรา', 'บอทแสงดาว', 'บอทเพชร', 'บอทโชคดี'
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

// --- Visual Animation / Visual FX Layer ---
let prevGStateSnapshot = null;

function getFXLayer() {
	return document.getElementById('visual-fx-layer');
}

function getCenterCoords(elem) {
	if (!elem) return { x: window.innerWidth / 2, y: window.innerHeight / 2 };
	const rect = elem.getBoundingClientRect();
	return {
		x: rect.left + rect.width / 2,
		y: rect.top + rect.height / 2
	};
}

function spawnFlyingCard(fromElem, toElem, cardData = null) {
	const fxLayer = getFXLayer();
	if (!fxLayer) return;

	const fromBox = getCenterCoords(fromElem);
	const toBox = getCenterCoords(toElem);

	const flyingCard = document.createElement('div');
	flyingCard.setAttribute('aria-hidden', 'true');

	if (cardData) {
		const isRed = (cardData.suit === '♥' || cardData.suit === '♦');
		flyingCard.className = `crazy-card visual-fx-card ${isRed ? 'red-suit' : 'black-suit'}`;
		flyingCard.innerHTML = `<div class="card-top">${cardData.rank}</div><div class="card-center">${cardData.suit}</div><div class="card-bottom">${cardData.rank}</div>`;
	} else {
		flyingCard.className = 'crazy-card visual-fx-card';
		flyingCard.style.background = 'linear-gradient(135deg, #1e293b 0%, #0f172a 100%)';
		flyingCard.style.border = '2px solid #f8fafc';
		flyingCard.innerHTML = `<div style="display:flex;height:100%;align-items:center;justify-content:center;color:#38bdf8;font-weight:900;font-size:10px;">CRAZY</div>`;
	}

	flyingCard.style.position = 'fixed';
	flyingCard.style.left = `${fromBox.x - 30}px`;
	flyingCard.style.top = `${fromBox.y - 45}px`;
	flyingCard.style.width = '60px';
	flyingCard.style.height = '90px';
	flyingCard.style.zIndex = '9999';
	flyingCard.style.pointerEvents = 'none';
	flyingCard.style.transition = 'all 0.35s cubic-bezier(0.25, 1, 0.5, 1)';
	flyingCard.style.transform = 'scale(0.8) rotate(-10deg)';
	flyingCard.style.opacity = '0.95';

	fxLayer.appendChild(flyingCard);

	requestAnimationFrame(() => {
		flyingCard.style.left = `${toBox.x - 30}px`;
		flyingCard.style.top = `${toBox.y - 45}px`;
		flyingCard.style.transform = 'scale(1) rotate(0deg)';
		flyingCard.style.opacity = '1';
	});

	setTimeout(() => {
		if (flyingCard.parentNode) {
			flyingCard.parentNode.removeChild(flyingCard);
		}
	}, 400);
}

function spawnSuitChangeFX(suit) {
	const fxLayer = getFXLayer();
	const centerElem = document.getElementById('board-center-container');
	if (!fxLayer || !centerElem) return;

	const coords = getCenterCoords(centerElem);
	const suitEl = document.createElement('div');
	suitEl.setAttribute('aria-hidden', 'true');
	const isRed = (suit === '♥' || suit === '♦');

	suitEl.style.position = 'fixed';
	suitEl.style.left = `${coords.x - 35}px`;
	suitEl.style.top = `${coords.y - 35}px`;
	suitEl.style.width = '70px';
	suitEl.style.height = '70px';
	suitEl.style.borderRadius = '50%';
	suitEl.style.display = 'flex';
	suitEl.style.alignItems = 'center';
	suitEl.style.justifyContent = 'center';
	suitEl.style.fontSize = '44px';
	suitEl.style.fontWeight = 'bold';
	suitEl.style.color = isRed ? '#f43f5e' : '#f8fafc';
	suitEl.style.background = 'rgba(15, 23, 42, 0.95)';
	suitEl.style.border = `3px solid ${isRed ? '#f43f5e' : '#fde047'}`;
	suitEl.style.boxShadow = `0 0 25px ${isRed ? 'rgba(244, 63, 94, 0.8)' : 'rgba(253, 224, 71, 0.8)'}`;
	suitEl.style.zIndex = '9999';
	suitEl.style.pointerEvents = 'none';
	suitEl.style.transition = 'all 0.5s cubic-bezier(0.175, 0.885, 0.32, 1.275)';
	suitEl.style.transform = 'scale(0.2)';
	suitEl.style.opacity = '0';
	suitEl.textContent = suit;

	fxLayer.appendChild(suitEl);

	requestAnimationFrame(() => {
		suitEl.style.transform = 'scale(1.3) translateY(-15px)';
		suitEl.style.opacity = '1';
	});

	setTimeout(() => {
		suitEl.style.opacity = '0';
		suitEl.style.transform = 'scale(1.6) translateY(-35px)';
		setTimeout(() => {
			if (suitEl.parentNode) suitEl.parentNode.removeChild(suitEl);
		}, 250);
	}, 550);
}

function spawnWinnerConfetti() {
	const fxLayer = getFXLayer();
	if (!fxLayer) return;

	const colors = ['#fde047', '#10b981', '#0ea5e9', '#f43f5e', '#a855f7', '#ffffff'];
	const count = 45;

	for (let i = 0; i < count; i++) {
		const particle = document.createElement('div');
		particle.setAttribute('aria-hidden', 'true');
		const color = colors[Math.floor(Math.random() * colors.length)];
		const startX = Math.random() * window.innerWidth;
		const startY = -20;
		const endX = startX + (Math.random() - 0.5) * 200;
		const endY = window.innerHeight + 50;
		const duration = 1800 + Math.random() * 1800;
		const size = 8 + Math.random() * 8;

		particle.style.position = 'fixed';
		particle.style.left = `${startX}px`;
		particle.style.top = `${startY}px`;
		particle.style.width = `${size}px`;
		particle.style.height = `${size * (Math.random() > 0.5 ? 1 : 2)}px`;
		particle.style.backgroundColor = color;
		particle.style.borderRadius = Math.random() > 0.5 ? '50%' : '2px';
		particle.style.zIndex = '9999';
		particle.style.pointerEvents = 'none';
		particle.style.transition = `transform ${duration}ms linear, opacity ${duration}ms linear`;
		particle.style.transform = `translate(0, 0) rotate(0deg)`;
		particle.style.opacity = '1';

		fxLayer.appendChild(particle);

		requestAnimationFrame(() => {
			particle.style.transform = `translate(${endX - startX}px, ${endY - startY}px) rotate(${Math.random() * 720}deg)`;
			particle.style.opacity = '0';
		});

		setTimeout(() => {
			if (particle.parentNode) particle.parentNode.removeChild(particle);
		}, duration);
	}
}

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

const btnLangLogin = document.getElementById('btn-lang-login');
const btnLangMain = document.getElementById('btn-lang-main');
if (btnLangLogin) btnLangLogin.onclick = toggleLanguage;
if (btnLangMain) btnLangMain.onclick = toggleLanguage;

const btnRefreshAudio = document.getElementById('btn-refresh-audio');
if (btnRefreshAudio) {
	btnRefreshAudio.onclick = () => {
		if (audioCtx.state === 'suspended') audioCtx.resume();
		playSound('select');
		announce(currentLang === 'en' ? 'Audio refreshed' : 'รีเฟรชเสียงแล้ว');
	};
}

const nameInput = document.getElementById('player-name-input');
const btnConfirmName = document.getElementById('btn-confirm-name');

nameInput.addEventListener('input', () => {
	if (nameInput.value.trim().length > 0) { btnConfirmName.disabled = false; }
	else { btnConfirmName.disabled = true; }
});

nameInput.addEventListener('keydown', (e) => {
	if ((e.key === 'Enter' || e.key === 'Return') && !btnConfirmName.disabled && nameInput.value.trim().length >= 1) {
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
	lastRenderedRooms = snapshot.val();
	if (document.getElementById('screen-main').classList.contains('active')) renderRoomList(lastRenderedRooms);
});

function renderRoomList(rooms) {
	lastRenderedRooms = rooms;
	const list = document.getElementById('room-list');
	const isEn = currentLang === 'en';
	const activeRooms = rooms ? Object.entries(rooms).filter(([id, r]) => Date.now() - r.lastActive < 120000 && r.status === 'waiting' && r.currentPlayers < MAX_PLAYERS) : [];
	if (activeRooms.length === 0) {
		list.innerHTML = `<li id="empty-room-msg" style="text-align: center; color: var(--text-muted);">${isEn ? 'No open rooms available' : 'ไม่มีห้องที่เปิดอยู่'}</li>`;
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
		const labelHtml = isEn
			? `<span>Room: ${id}</span> <span>${room.currentPlayers}/${MAX_PLAYERS} players</span>`
			: `<span>ห้อง: ${id}</span> <span>${room.currentPlayers}/${MAX_PLAYERS} คน</span>`;
		const ariaText = isEn
			? `Join room ${id}, ${room.currentPlayers} of ${MAX_PLAYERS} players`
			: `เข้าร่วมห้อง ${id} มีผู้เล่น ${room.currentPlayers} จาก ${MAX_PLAYERS} คน`;

		if (li) {
			const btn = li.querySelector('button');
			if (btn) {
				btn.innerHTML = labelHtml;
				btn.setAttribute('aria-label', ariaText);
			}
		} else {
			li = document.createElement('li'); 
			li.style.margin = '10px 0';
			li.setAttribute('data-room-id', id);
			const btn = document.createElement('button');
			btn.style.width = '100%'; btn.style.maxWidth = '100%'; btn.style.display = 'flex'; btn.style.justifyContent = 'space-between';
			btn.innerHTML = labelHtml;
			btn.setAttribute('aria-label', ariaText);
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
	if (!myPeerId) {
		announce(currentLang === 'en' ? 'System is initializing, please wait...' : 'ระบบกำลังเตรียมพร้อม กรุณารอสักครู่');
		return;
	}
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
		announce(currentLang === 'en'
			? 'Room is ready! Wait for friends or add bots right away.'
			: 'ห้องพร้อมแล้ว รอเพื่อนหรือเพิ่มบอทได้ทันที');
	} catch (err) {
		isCreatingRoom = false;
		document.getElementById('btn-create-room').disabled = false;
		announce(currentLang === 'en' ? 'Failed to create room' : 'สร้างห้องไม่สำเร็จ');
	}
};

function joinRoom(roomId, hostPeerId) {
	if (!myPeerId) {
		announce(currentLang === 'en' ? 'System is initializing, please wait...' : 'ระบบกำลังเตรียมพร้อม กรุณารอสักครู่');
		isJoiningRoom = false;
		return;
	}
	isHost = false; currentRoomId = roomId;
	announce(currentLang === 'en' ? 'Connecting to Host...' : 'กำลังเชื่อมต่อไปยัง Host...');
	hostConnection = peer.connect(hostPeerId, { reliable: true });
	hostConnection.on('error', () => {
		announce(currentLang === 'en' ? 'Connection failed' : 'การเชื่อมต่อล้มเหลว', true);
		leaveLobby();
	});
	hostConnection.on('open', () => {
		hostConnection.send({ type: 'joinReq', peerId: myPeerId, name: myName });
		enterLobby();
		hostConnection.on('data', handleClientData);
		hostConnection.on('close', () => {
			announce(currentLang === 'en' ? 'Host disconnected' : 'Host หลุดการเชื่อมต่อ', true);
			leaveLobby();
		});
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
	announce(currentLang === 'en' ? 'Left the room' : 'ออกจากห้องแล้ว');
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
	const isEn = currentLang === 'en';
	list.innerHTML = '';
	document.getElementById('lobby-player-count').textContent = players.length;
	players.forEach(p => {
		const li = document.createElement('li');
		li.style.padding = '12px 15px'; li.style.background = 'rgba(255,255,255,0.1)'; li.style.marginBottom = '8px'; li.style.borderRadius = '12px';
		li.style.display = 'flex'; li.style.alignItems = 'center'; li.style.gap = '10px'; li.style.border = '1px solid rgba(255,255,255,0.05)';
		const dispName = formatPlayerDisplayName(p.name);
		const youTag = p.id === myPeerId ? (isEn ? '(You)' : '(คุณ)') : '';
		li.innerHTML = `<span aria-hidden="true" style="font-size: 20px;">${p.isBot?'🤖':'👤'}</span> <strong>${dispName}</strong> ${youTag}`;
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
		p.isBot = true; 
		if (!p.name.startsWith('บอท')) {
			p.name = `บอท${p.name}`;
		}
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
function getSuitName(suit) {
	const isEn = currentLang === 'en';
	if (suit === '♥' || suit === 'โพแดง') return isEn ? 'Hearts' : 'โพแดง';
	if (suit === '♦' || suit === 'ข้าวหลามตัด') return isEn ? 'Diamonds' : 'ข้าวหลามตัด';
	if (suit === '♣' || suit === 'ดอกจิก') return isEn ? 'Clubs' : 'ดอกจิก';
	if (suit === '♠' || suit === 'โพดำ') return isEn ? 'Spades' : 'โพดำ';
	return suit;
}

function getSuitNameThai(suit) {
	if (suit === '♥' || suit === 'โพแดง') return 'โพแดง';
	if (suit === '♦' || suit === 'ข้าวหลามตัด') return 'ข้าวหลามตัด';
	if (suit === '♣' || suit === 'ดอกจิก') return 'ดอกจิก';
	if (suit === '♠' || suit === 'โพดำ') return 'โพดำ';
	return suit;
}

function generateDeck() {
	const deck = [];
    const suits = ['♥', '♦', '♣', '♠'];
    const ranks = ['A', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K'];
    suits.forEach(suit => {
        ranks.forEach(rank => {
            let suitName = getSuitNameThai(suit);
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
	const isEn = currentLang === 'en';
	const dispCardName = formatCardNameForLang(card);
	btn.className = `crazy-card ${isRed ? 'red-suit' : 'black-suit'} ${(index >= 0 && !playable && !isTop) ? 'disabled' : ''}`;
	btn.innerHTML = `<div class="card-top">${card.rank}</div><div class="card-center">${card.suit}</div><div class="card-bottom">${card.rank}</div>`;
	
	if (index >= 0 && !isTop) {
		const statusStr = isEn ? (playable ? 'Playable' : 'Unplayable') : (playable ? 'ลงได้' : 'ลงไม่ได้');
		btn.setAttribute('aria-label', `${dispCardName} ${statusStr}`);
		if (playable) {
			btn.onclick = (e) => {
                if (card.rank === '8') {
                    // เปิด Modal เลือกดอก
                    pendingPlayIndex = index;
                    const modal = document.getElementById('suit-picker-modal');
                    if (modal) {
                        modal.style.display = 'flex';
                        const modalHeading = modal.querySelector('h1, h2, h3, [role="heading"]');
                        if (modalHeading) {
                            if (!modalHeading.hasAttribute('tabindex')) {
                                modalHeading.setAttribute('tabindex', '-1');
                            }
                            modalHeading.focus();
                        }
                    }
                    announce(currentLang === 'en' ? "Please choose a suit to change to" : "กรุณาเลือกเปลี่ยนดอกการ์ด");
                } else {
				    sendAction('play', { index });
                }
			};
		} else {
			btn.setAttribute('aria-disabled', 'true');
		}
	} else if (isTop) {
        btn.setAttribute('aria-label', isEn ? `Discard pile card is ${dispCardName}` : `การ์ดกองทิ้งคือ ${dispCardName}`);
    }
	return btn;
}

let pendingPlayIndex = -1;

document.querySelectorAll('.suit-btn').forEach(btn => {
    btn.onclick = (e) => {
        const selectedSuit = e.currentTarget.getAttribute('data-suit') || e.target.getAttribute('data-suit');
        const modal = document.getElementById('suit-picker-modal');
        if (modal) {
            modal.style.display = 'none';
        }
        
        sendAction('play', { index: pendingPlayIndex, activeSuit: selectedSuit });
        pendingPlayIndex = -1;

        const focusTarget = document.getElementById('top-status-bar') || document.getElementById('screen-game');
        if (focusTarget) {
            if (!focusTarget.hasAttribute('tabindex')) {
                focusTarget.setAttribute('tabindex', '-1');
            }
            focusTarget.focus();
        }
    };
});

// เพิ่มปุ่มยกเลิกการลงการ์ด 8
const modalContent = document.querySelector('#suit-picker-modal .modal-content');
if (modalContent) {
    const cancelBtn = document.createElement('button');
    cancelBtn.style.cssText = 'font-size: 24px; padding: 15px; color: white; background: #64748b; border: 2px solid rgba(255,255,255,0.2); border-radius: 8px; cursor: pointer;';
    cancelBtn.textContent = currentLang === 'en' ? 'Cancel' : 'ยกเลิก';
    cancelBtn.setAttribute('aria-label', currentLang === 'en' ? 'Cancel playing 8 card' : 'ยกเลิกการลงการ์ด 8');
    cancelBtn.onclick = () => {
        const modal = document.getElementById('suit-picker-modal');
        if (modal) {
            modal.style.display = 'none';
        }
        pendingPlayIndex = -1;
        announce(currentLang === 'en' ? 'Cancelled playing 8 card' : 'ยกเลิกการลงการ์ด 8', true);
        
        const focusTarget = document.getElementById('my-cards-container');
        if (focusTarget) {
            if (!focusTarget.hasAttribute('tabindex')) {
                focusTarget.setAttribute('tabindex', '-1');
            }
            focusTarget.focus();
        }
    };
    modalContent.appendChild(cancelBtn);
    suitModalCancelBtn = cancelBtn;
}

applyStaticLanguage();

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
	animDiv.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:radial-gradient(circle at center, #1b382b 0%, #000000 90%);z-index:10000;display:flex;justify-content:center;align-items:center;color:#fff;font-size:5rem;font-weight:900;text-shadow:0 0 30px #4CAF50,0 0 60px #2e7d32; flex-direction:column;transition:all 0.3s ease;';
	document.body.appendChild(animDiv);
	
	const steps = [
		{ time: 200, text: 'CRAZY EIGHTS' }, { time: 1200, text: '3' }, { time: 2200, text: '2' }, { time: 3200, text: '1' }, { time: 4200, text: 'Enjoy' }
	];
	steps.forEach(step => {
		setTimeout(() => { 
			animDiv.textContent = step.text; 
			animDiv.style.transform = 'scale(1.1)';
			setTimeout(() => { animDiv.style.transform = 'scale(1)'; }, 150);
			announce(step.text, true); 
		}, step.time);
	});
	setTimeout(() => { animDiv.remove(); if(callback) callback(); }, 5000);
}

function broadcastTurnStart() {
	if (!isHost) return;
	const currentPlayer = players[game.turnIndex];
    let topCard = game.discardPile[game.discardPile.length - 1];
    let cardName = topCard ? topCard.name : '';
    
    if (topCard && game.activeSuit && (topCard.rank === '8' || game.activeSuit !== topCard.suit)) {
        cardName = `${topCard.name} เปลี่ยนดอกเป็น ${getSuitNameThai(game.activeSuit)}`;
    }
    
    let msg = `ถึงรอบของ ${currentPlayer.name} การ์ดกองทิ้งคือ ${cardName}`;
	broadcastAnnounce(msg);
}

function initGame() {
	game.status = 'playing'; game.matchOver = false;
	game.deck = generateDeck();
    game.discardPile = [];
    game.activeSuit = null;
    previousTurnIndex = -1;

    // Randomize Turn Sequence
    game.turnOrder = [];
    for(let i = 0; i < players.length; i++) {
        game.turnOrder.push(i);
    }
    game.turnOrder.sort(() => Math.random() - 0.5);
    game.turnSequenceIndex = 0;
    game.turnIndex = game.turnOrder[game.turnSequenceIndex];
	
	players.forEach((p, i) => {
		game.playerStates[p.id] = { hand: [] };
		for(let k=0; k<7; k++) game.playerStates[p.id].hand.push(game.deck.pop());
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
        turnOrder: game.turnOrder,
        turnSequenceIndex: game.turnSequenceIndex,
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
	else if (data.type === 'announce') {
		const localMsg = getLocalizedMessage(data.message);
		announce(localMsg, data.assertive);
		logEvent(localMsg);
	}
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
	lastRenderedGameState = gState;
	const isEn = currentLang === 'en';

	// Visual FX trigger on real state changes
	if (prevGStateSnapshot && gState.status === 'playing' && prevGStateSnapshot.status === 'playing') {
		// 1. Card Draw FX
		if (gState.deckCount < prevGStateSnapshot.deckCount) {
			let drawerId = null;
			if (gState.players && prevGStateSnapshot.playerStates) {
				for (let p of gState.players) {
					const prevCount = prevGStateSnapshot.playerStates[p.id]?.cardCount || 0;
					const currCount = gState.playerStates[p.id]?.cardCount || 0;
					if (currCount > prevCount) {
						drawerId = p.id;
						break;
					}
				}
			}
			const deckEl = document.getElementById('ui-deck-pile');
			let targetEl = null;
			if (drawerId === myPeerId) {
				targetEl = document.getElementById('my-cards-container');
			} else if (drawerId) {
				const statusBar = document.getElementById('top-status-bar');
				if (statusBar && gState.turnOrder) {
					const idxInOrder = gState.turnOrder.findIndex(i => gState.players[i] && gState.players[i].id === drawerId);
					if (idxInOrder !== -1 && statusBar.children[idxInOrder]) {
						targetEl = statusBar.children[idxInOrder];
					}
				}
			}
			if (deckEl && targetEl) {
				spawnFlyingCard(deckEl, targetEl, null);
			}
		}

		// 2. Card Play FX
		if (gState.topCard && prevGStateSnapshot.topCard && gState.topCard.id !== prevGStateSnapshot.topCard.id) {
			let playerIdx = prevGStateSnapshot.turnIndex;
			let player = gState.players[playerIdx];
			let fromEl = null;
			if (player && player.id === myPeerId) {
				fromEl = document.getElementById('my-cards-container');
			} else if (player) {
				const statusBar = document.getElementById('top-status-bar');
				if (statusBar && gState.turnOrder) {
					const idxInOrder = gState.turnOrder.findIndex(i => i === playerIdx);
					if (idxInOrder !== -1 && statusBar.children[idxInOrder]) {
						fromEl = statusBar.children[idxInOrder];
					}
				}
			}
			const centerCont = document.getElementById('board-center-container');
			if (fromEl && centerCont) {
				spawnFlyingCard(fromEl, centerCont, gState.topCard);
			}
		}

		// 3. 8 Suit Change FX
		if (gState.activeSuit && prevGStateSnapshot.activeSuit && gState.activeSuit !== prevGStateSnapshot.activeSuit) {
			spawnSuitChangeFX(gState.activeSuit);
		}
	}

	// Update snapshot
	prevGStateSnapshot = {
		status: gState.status,
		deckCount: gState.deckCount,
		turnIndex: gState.turnIndex,
		activeSuit: gState.activeSuit,
		topCard: gState.topCard ? { ...gState.topCard } : null,
		playerStates: {}
	};
	if (gState.players) {
		gState.players.forEach(p => {
			prevGStateSnapshot.playerStates[p.id] = { cardCount: gState.playerStates[p.id]?.cardCount || 0 };
		});
	}

	const turnPlayer = gState.players[gState.turnIndex];
	if (turnPlayer) {
		const headingEl = document.getElementById('current-turn-heading');
		const dispTurnName = formatPlayerDisplayName(turnPlayer.name);
		headingEl.textContent = isEn ? `${dispTurnName}'s Turn` : `รอบของ ${dispTurnName}`;
		headingEl.style.color = '#ffffff';
		if(turnPlayer.id === myPeerId) {
			headingEl.style.background = 'rgba(46, 125, 50, 0.85)';
			headingEl.style.borderColor = 'var(--primary)';
		} else {
			headingEl.style.background = 'rgba(15, 23, 42, 0.85)';
			headingEl.style.borderColor = 'rgba(255, 255, 255, 0.2)';
		}
	}

	if (gState.status === 'playing') {
		if (turnPlayer && turnPlayer.id === myPeerId && gState.turnIndex !== previousTurnIndex) {
			playSound('abc');
		}
		previousTurnIndex = gState.turnIndex;
	} else {
		previousTurnIndex = -1;
	}

	const statusBar = document.getElementById('top-status-bar');
	statusBar.innerHTML = '';
	let ariaStatusBarText = ``;

	const mascots = ['🐱', '🐶', '🐰', '🦊', '🐼', '🐸'];
	const displayOrder = (gState.turnOrder && gState.turnOrder.length === gState.players.length) 
		? gState.turnOrder 
		: gState.players.map((_, i) => i);

	displayOrder.forEach((idx) => {
		const p = gState.players[idx];
		if (!p) return;
		const pState = gState.playerStates[p.id];
		const isTurn = idx === gState.turnIndex;
		const dispPName = formatPlayerDisplayName(p.name);
		ariaStatusBarText += isEn
			? `${dispPName} has ${pState.cardCount} cards. `
			: `${dispPName} มี ${pState.cardCount} ใบ. `;
		const charDiv = document.createElement('div');
		charDiv.className = `character-card ${isTurn ? 'is-turn' : ''}`; charDiv.setAttribute('aria-hidden', 'true');
		const mascot = mascots[idx % 6];
		const countVisualText = isEn ? `${pState.cardCount} left` : `เหลือ ${pState.cardCount} ใบ`;
		charDiv.innerHTML = `<div class="char-cards-count">${countVisualText}</div><div class="char-name">${mascot} ${dispPName}</div>`;
		statusBar.appendChild(charDiv);
	});
	statusBar.setAttribute('aria-label', ariaStatusBarText);

	document.getElementById('deck-count-visual').textContent = gState.deckCount;
	document.getElementById('deck-aria-label').textContent = isEn
		? `Draw pile has ${gState.deckCount} cards left`
		: `กองจั่วเหลือ ${gState.deckCount} ใบ`;

	const centerCont = document.getElementById('board-center-container');
	const centerAria = document.getElementById('board-center-aria');
	centerCont.innerHTML = '';
    
    if (gState.topCard) {
        const displayCard = { ...gState.topCard };
        if (gState.activeSuit && gState.activeSuit !== gState.topCard.suit) {
            displayCard.suit = gState.activeSuit;
        }
        const activeSuitName = getSuitName(gState.activeSuit);
        const origCardName = formatCardNameForLang(gState.topCard);
        if (gState.topCard.rank === '8' || (gState.activeSuit && gState.activeSuit !== gState.topCard.suit)) {
            displayCard.name = isEn
                ? `${origCardName} (Suit changed to ${activeSuitName})`
                : `${gState.topCard.name} เปลี่ยนดอกเป็น ${activeSuitName}`;
            centerAria.textContent = isEn
                ? `Discard pile: ${origCardName}, suit changed to ${activeSuitName}`
                : `กองทิ้ง: ${gState.topCard.name} เปลี่ยนดอกเป็น ${activeSuitName}`;
        } else {
            centerAria.textContent = isEn
                ? `Discard pile: ${origCardName}`
                : `กองทิ้ง: ${gState.topCard.name}`;
        }
        centerCont.appendChild(renderCardHTML(displayCard, -1, false, true));
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
		if (game.deck.length === 0) return;
		const card = game.deck.pop();
		game.playerStates[peerId].hand.push(card);
		broadcastSound('jua');

		const publicMsg = `${currentPlayer.name} จั่วการ์ด 1 ใบ`;
		const privateMsg = `คุณจั่วได้การ์ด ${card.name}`;

        const topCard = game.discardPile[game.discardPile.length - 1];
        const isPlayable = canPlayCard(card, game.activeSuit, topCard);

		connections.forEach(c => {
			if (c.open) {
				const targetId = c.customPeerId || c.peer;
				if (targetId === peerId) {
					c.send({ type: 'announce', message: privateMsg, assertive: true });
					if (!isPlayable) c.send({ type: 'announce', message: 'ไม่มีการ์ดลงได้ จบตา', assertive: true });
				} else {
					c.send({ type: 'announce', message: publicMsg, assertive: false });
					if (!isPlayable) c.send({ type: 'announce', message: 'ไม่มีการ์ดลงได้ จบตา', assertive: false });
				}
			}
		});

		if (myPeerId === peerId) {
			const locPriv = getLocalizedMessage(privateMsg);
			announce(locPriv, true);
			logEvent(locPriv);
			if (!isPlayable) {
				const locNoPlay = getLocalizedMessage('ไม่มีการ์ดลงได้ จบตา');
				announce(locNoPlay, true);
				logEvent(locNoPlay);
			}
		} else {
			const locPub = getLocalizedMessage(publicMsg);
			announce(locPub, false);
			logEvent(locPub);
			if (!isPlayable) {
				const locNoPlay = getLocalizedMessage('ไม่มีการ์ดลงได้ จบตา');
				announce(locNoPlay, false);
				logEvent(locNoPlay);
			}
		}

		broadcastGameState();
        
        if (isPlayable) {
            if (currentPlayer.isBot) {
                setTimeout(() => {
                    let playPayload = { index: game.playerStates[peerId].hand.length - 1 };
                    if (card.rank === '8') {
                        const suits = ['♥', '♦', '♣', '♠'];
                        playPayload.activeSuit = suits[Math.floor(Math.random() * suits.length)];
                    }
                    handlePlayerAction(peerId, 'play', playPayload);
                }, 1500);
            }
        } else {
            setTimeout(() => advanceTurn(), 1500);
        }
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
        broadcastAnnounce(`${currentPlayer.name} ลง ${card.name} และเปลี่ยนดอกเป็น ${getSuitNameThai(game.activeSuit)}`);
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
        broadcastAnnounce('กองจั่วหมด สับการ์ดใหม่เรียบร้อยแล้ว');
    }
}

function advanceTurn() {
    reshuffleDeck();
    if (game.turnOrder && game.turnOrder.length === players.length) {
        game.turnSequenceIndex = (game.turnSequenceIndex + game.direction + players.length) % players.length;
        game.turnIndex = game.turnOrder[game.turnSequenceIndex];
    } else {
        game.turnIndex = (game.turnIndex + game.direction + players.length) % players.length;
    }
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
			if (game.status === 'playing' && players[game.turnIndex].id === currentPlayer.id && !currentPlayer.isBot && !state.isProcessing) {
                const topCard = game.discardPile[game.discardPile.length - 1];
                let playableIndex = -1;
                
                for(let i=0; i<state.hand.length; i++) {
                    if(canPlayCard(state.hand[i], game.activeSuit, topCard)) {
                        playableIndex = i;
                        break;
                    }
                }
                
                if (playableIndex !== -1) {
                    const card = state.hand[playableIndex];
                    let payload = { index: playableIndex };
                    if (card.rank === '8') {
                        const suits = ['♥', '♦', '♣', '♠'];
                        payload.activeSuit = suits[Math.floor(Math.random() * suits.length)];
                    }
                    handlePlayerAction(currentPlayer.id, 'play', payload);
                } else {
                    if (game.deck.length > 0) {
                        handlePlayerAction(currentPlayer.id, 'draw');
                    } else {
                        advanceTurn();
                    }
                }
			}
		}, 32000);
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
	lastDisplayedWinner = winnerName;
	lastDisplayedStats = resultStats;
	switchScreen('screen-result', 'title-result');
	stopBGM();
	const isEn = currentLang === 'en';
	let winnerStat = resultStats.find(r => r.isWinner) || resultStats[0];
	let losers = resultStats.filter(r => !r.isWinner);
	const dispWinnerName = formatPlayerDisplayName(winnerName);
	let html = isEn
		? `<div style="color: var(--focus-ring); margin-bottom: 10px;">👑 ${dispWinnerName} is the winner (0 cards left)</div>`
		: `<div style="color: var(--focus-ring); margin-bottom: 10px;">👑 ${dispWinnerName} เป็นผู้ชนะ (การ์ดหมดมือ)</div>`;
	losers.forEach(l => {
		const dispLoserName = formatPlayerDisplayName(l.name);
		html += isEn
			? `<div style="color: #ffcdd2; margin-bottom: 10px;">❌ ${dispLoserName} has ${l.cardsLeft} card(s) left</div>`
			: `<div style="color: #ffcdd2; margin-bottom: 10px;">❌ ${dispLoserName} เหลือการ์ด ${l.cardsLeft} ใบ</div>`;
	});
	document.getElementById('winner-text').innerHTML = html;
	document.getElementById('result-status-bar').innerHTML = '';
	spawnWinnerConfetti();
}

// --- Keyboard Shortcuts ---
document.addEventListener('keydown', (e) => {
    const activeTag = document.activeElement ? document.activeElement.tagName.toLowerCase() : '';
    if (activeTag === 'input' || activeTag === 'textarea') return;

    if (e.altKey) {
        const key = e.key.toLowerCase();
        if (key === 'p') {
            e.preventDefault();
            const btnDraw = document.getElementById('btn-draw');
            if (btnDraw && !btnDraw.disabled) {
                btnDraw.click();
            }
        } else if (key === 'c') {
            e.preventDefault();
            const centerAria = document.getElementById('board-center-aria');
            if (centerAria && centerAria.textContent) {
                announce(centerAria.textContent, true);
            }
        } else if (key === 'k') {
            e.preventDefault();
            const deckAria = document.getElementById('deck-aria-label');
            if (deckAria && deckAria.textContent) {
                announce(deckAria.textContent, true);
            }
        } else if (key === 'a') {
            e.preventDefault();
            const statusBar = document.getElementById('top-status-bar');
            if (statusBar) {
                const label = statusBar.getAttribute('aria-label');
                if (label) announce(label, true);
            }
        }
    }
});