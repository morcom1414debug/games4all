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

// ==================== Bilingual Support (TH / EN) ====================
let currentLang = localStorage.getItem('domino_lang') === 'en' ? 'en' : 'th';

const DOMINO_I18N = {
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
		rules_title: "คู่มือการเล่น DOMINO",
		rules_content_aria: "เนื้อหาคู่มือ",
		rules_html: `
            <p style="font-size: 16px; margin-bottom: 15px;"><strong>🎯 เป้าหมาย:</strong> วางโดมิโน่ให้หมดมือเป็นคนแรก เพื่อเป็นผู้ชนะในรอบนั้น</p>
            <h3 style="font-size: 16px; margin-top: 15px; text-align: left; border-left: 4px solid var(--primary); padding-left: 10px;">กติกาการวาง</h3>
            <ul style="margin-left: 20px; margin-bottom: 15px; color: var(--text-muted);">
                <li>โดมิโน่จะวางต่อกันที่ <strong style="color:#fff;">ปลายซ้าย</strong> หรือ <strong style="color:#fff;">ปลายขวา</strong></li>
                <li>เลขด้านใดด้านหนึ่งของโดมิโน่ต้องตรงกับปลายที่จะวาง</li>
                <li>หากไม่มีตัววางได้ ต้อง <strong style="color:var(--info);">กดปุ่มจั่ว</strong> (สูงสุด 3 ตัวต่อ Turn)</li>
                <li>เมื่อครบ 3 ตัวแล้วยังไม่มีตัววางได้ ระบบจะ <strong style="color:var(--danger);">จบ Turn อัตโนมัติ</strong></li>
                <li>เมื่อวางโดมิโน่แล้ว จะ <strong style="color:#fff;">จบ Turn ทันที</strong> (เล่นได้เพียง 1 ตัวต่อ Turn) ยกเว้นมีตัว Double วางต่อได้</li>
            </ul>
            <h3 style="font-size: 16px; margin-top: 15px; text-align: left; border-left: 4px solid var(--primary); padding-left: 10px;">โดมิโน่พิเศษ</h3>
            <ul style="margin-left: 20px; margin-bottom: 15px; color: var(--text-muted);">
                <li><strong style="color:var(--sp-sleep);">หลับ:</strong> ข้ามตาผู้เล่นคนถัดไป</li>
                <li><strong style="color:var(--sp-draw2);">2+:</strong> ผู้เล่นถัดไปจั่ว 2 ตัว และถูกข้ามตา</li>
                <li><strong style="color:var(--sp-draw3);">3+:</strong> ผู้เล่นถัดไปจั่ว 3 ตัว และถูกข้ามตา</li>
                <li><strong style="color:var(--sp-reverse);">ย้อนศร:</strong> กลับทิศทางการเล่น</li>
            </ul>
            <h3 style="font-size: 16px; margin-top: 15px; text-align: left; border-left: 4px solid var(--primary); padding-left: 10px;">กฎ DOMINO!</h3>
            <ul style="margin-left: 20px; margin-bottom: 10px; color: var(--text-muted);">
                <li>เมื่อเหลือโดมิโน่ 2 ตัว และถึงตาที่สามารถวางได้ ต้องกดปุ่ม <strong style="color:var(--danger);">DOMINO!</strong> ก่อนวาง</li>
                <li>หากลืมกดและวางไปจนเหลือ 1 ตัว จะถูกลงโทษจั่ว 2 ตัว</li>
            </ul>
            <h3 style="font-size: 16px; margin-top: 15px; text-align: left; border-left: 4px solid var(--primary); padding-left: 10px;">แป้นพิมพ์ลัด</h3>
            <ul>
              <li><strong>Alt + C</strong>: อ่านสถานะปลายกระดานซ้ายและขวา</li>
              <li><strong>Alt + K</strong>: อ่านจำนวนโดมิโน่ในกองจั่ว</li>
              <li><strong>Alt + P</strong>: กดปุ่มจั่วโดมิโน่</li>
              <li><strong>Alt + O</strong>: กดปุ่มประกาศ DOMINO!</li>
              <li><strong>Alt + A</strong>: อ่านสถานะจำนวนโดมิโน่ในมือของผู้เล่นทุกคน</li>
            </ul>
		`,
		btn_back_main: "กลับหน้าหลัก",
		lobby_prefix: "ล็อบบี้:",
		lobby_ready_msg: "เล่นได้สูงสุดที่ 6 คนต่อห้อง เจ้าของห้องพร้อมก็กดเริ่มเกมได้นะ เพื่อสนุกกัน",
		players_in_room_prefix: "ผู้เล่นในห้อง",
		bot_management_heading: "การจัดการ Bot",
		btn_remove_bot_aria: "ลดบอท",
		btn_add_bot_aria: "เพิ่มบอท",
		btn_start_game: "🚀 เริ่มเกม",
		btn_leave_lobby: "ออกจากห้อง",
		btn_refresh_audio: "🔊 รีเฟรช",
		btn_refresh_audio_aria: "รีเฟรชเสียง",
		board_left: "ซ้าย",
		board_right: "ขวา",
		my_cards_aria: "โดมิโน่ในมือ",
		btn_draw: "จั่ว",
		side_picker_title: "เลือกปลายที่จะวาง",
		btn_play_left: "วางปลายซ้าย",
		btn_play_right: "วางปลายขวา",
		btn_cancel: "ยกเลิก",
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
		rules_title: "DOMINO Game Manual",
		rules_content_aria: "Game manual content",
		rules_html: `
            <p style="font-size: 16px; margin-bottom: 15px;"><strong>🎯 Objective:</strong> Be the first player to play all dominoes in your hand to win the round.</p>
            <h3 style="font-size: 16px; margin-top: 15px; text-align: left; border-left: 4px solid var(--primary); padding-left: 10px;">Placement Rules</h3>
            <ul style="margin-left: 20px; margin-bottom: 15px; color: var(--text-muted);">
                <li>Dominoes are placed at either the <strong style="color:#fff;">Left End</strong> or the <strong style="color:#fff;">Right End</strong>.</li>
                <li>One side of your domino must match the value on the chosen board end.</li>
                <li>If you have no playable dominoes, press <strong style="color:var(--info);">Draw</strong> (up to 3 draws per turn).</li>
                <li>After 3 draws with no playable tile, your turn <strong style="color:var(--danger);">ends automatically</strong>.</li>
                <li>Playing a domino <strong style="color:#fff;">ends your turn immediately</strong> (1 tile per turn), unless you play a Double and can chain another tile on that same end.</li>
            </ul>
            <h3 style="font-size: 16px; margin-top: 15px; text-align: left; border-left: 4px solid var(--primary); padding-left: 10px;">Special Dominoes</h3>
            <ul style="margin-left: 20px; margin-bottom: 15px; color: var(--text-muted);">
                <li><strong style="color:var(--sp-sleep);">Sleep:</strong> Skips the next player's turn.</li>
                <li><strong style="color:var(--sp-draw2);">2+:</strong> Next player draws 2 tiles and is skipped.</li>
                <li><strong style="color:var(--sp-draw3);">3+:</strong> Next player draws 3 tiles and is skipped.</li>
                <li><strong style="color:var(--sp-reverse);">Reverse:</strong> Reverses the direction of play.</li>
            </ul>
            <h3 style="font-size: 16px; margin-top: 15px; text-align: left; border-left: 4px solid var(--primary); padding-left: 10px;">DOMINO! Rule</h3>
            <ul style="margin-left: 20px; margin-bottom: 10px; color: var(--text-muted);">
                <li>When you have 2 dominoes left and can play one, you must press <strong style="color:var(--danger);">DOMINO!</strong> before placing your tile.</li>
                <li>If you forget to press it and drop to 1 tile, you receive a 2-tile draw penalty.</li>
            </ul>
            <h3 style="font-size: 16px; margin-top: 15px; text-align: left; border-left: 4px solid var(--primary); padding-left: 10px;">Keyboard Shortcuts</h3>
            <ul>
              <li><strong>Alt + C</strong>: Announce left and right board ends</li>
              <li><strong>Alt + K</strong>: Announce remaining tiles in the draw pile</li>
              <li><strong>Alt + P</strong>: Press the Draw button</li>
              <li><strong>Alt + O</strong>: Press the DOMINO! button</li>
              <li><strong>Alt + A</strong>: Announce tile counts for all players</li>
            </ul>
		`,
		btn_back_main: "Back to Home",
		lobby_prefix: "Lobby:",
		lobby_ready_msg: "Up to 6 players per room. Host can press Start Game whenever ready!",
		players_in_room_prefix: "Players in Room",
		bot_management_heading: "Bot Management",
		btn_remove_bot_aria: "Remove bot",
		btn_add_bot_aria: "Add bot",
		btn_start_game: "🚀 Start Game",
		btn_leave_lobby: "Leave Room",
		btn_refresh_audio: "🔊 Refresh",
		btn_refresh_audio_aria: "Refresh audio",
		board_left: "Left",
		board_right: "Right",
		my_cards_aria: "Dominoes in hand",
		btn_draw: "Draw",
		side_picker_title: "Choose Board End",
		btn_play_left: "Play Left End",
		btn_play_right: "Play Right End",
		btn_cancel: "Cancel",
		result_title: "Match Results"
	}
};

const BOT_NAME_EN_MAP = {
	'บอทซุปเปอร์แมน': 'Bot Superman',
	'บอทแมงมุม': 'Bot Spider',
	'บอทมาริโอ้': 'Bot Mario',
	'บอทนารูโตะ': 'Bot Naruto',
	'บอทโซนิค': 'Bot Sonic',
	'บอทปิกาจู': 'Bot Pikachu',
	'บอทเบ็นเท็น': 'Bot Ben10',
	'บอทโดราเอมอน': 'Bot Doraemon',
	'บอทลูฟี่': 'Bot Luffy',
	'บอทคิริโตะ': 'Bot Kirito',
	'บอทไอรอนแมน': 'Bot IronMan',
	'บอทแบทแมน': 'Bot Batman',
	'บอทกัปตัน': 'Bot Captain',
	'บอทชินจัง': 'Bot Shinchan',
	'บอทโคนัน': 'Bot Conan',
	'บอทอุลตร้าแมน': 'Bot Ultraman',
	'บอทก็อดซิลล่า': 'Bot Godzilla',
	'บอทเซเลอร์มูน': 'Bot SailorMoon',
	'บอทซากุระ': 'Bot Sakura',
	'บอทดราก้อนบอล': 'Bot DragonBall'
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

function translateSpecialWordsToEn(text) {
	return String(text)
		.replace(/หลับ/g, 'Sleep')
		.replace(/ย้อนศร/g, 'Reverse')
		.replace(/\bsleep\b/gi, 'Sleep')
		.replace(/\breverse\b/gi, 'Reverse')
		.replace(/\bdraw2\b/gi, '2+')
		.replace(/\bdraw3\b/gi, '3+');
}

function translateAnnouncementToEn(rawMsg) {
	if (!rawMsg) return '';
	let s = replaceAllBotNamesToEn(rawMsg);

	// 1. ข้อความแจ้งเริ่มรอบ (broadcastTurnStart)
	const turnMatch = s.match(/^ถึงรอบของ (.+?) ขณะนี้เล่น(ตามเข็มนาฬิกา|ทวนเข็มนาฬิกา) (.*)$/);
	if (turnMatch) {
		const pName = turnMatch[1];
		const dirEn = turnMatch[2] === 'ตามเข็มนาฬิกา' ? 'clockwise' : 'counter-clockwise';
		let rest = turnMatch[3];

		rest = rest.replace('กระดานยังว่าง คุณสามารถเริ่มวางได้', 'The board is empty; you may place any tile.');
		rest = rest.replace(
			/ปลายซ้ายคือ (.+?) ปลายขวาคือ (.+?)(?=\s*\(|\s*ลงได้|\s*ลงไม่ได้|$)/,
			(_, lVal, rVal) => `Left end is ${translateSpecialWordsToEn(lVal)}, right end is ${translateSpecialWordsToEn(rVal)}.`
		);
		rest = rest.replace(
			/\(มีการล็อกฝั่งเล่นต่อที่ปลาย(ซ้าย|ขวา)\)/,
			(_, side) => `(Chain locked on the ${side === 'ซ้าย' ? 'left' : 'right'} end)`
		);
		rest = rest.replace(/ลงได้ (\d+) ตัว เลือกเลย/, '$1 playable tile(s). Choose one!');
		rest = rest.replace('ลงไม่ได้จั่วเลย', 'No playable tiles. Draw!');

		return `${pName}'s turn. Direction: ${dirEn}. ${rest}`.trim();
	}

	// 2. ข้อความล็อบบี้และการเชื่อมต่อ
	s = s.replace(/^(.+?) หลุดการเชื่อมต่อ เปลี่ยนเป็นบอทแล้ว$/, '$1 disconnected and was replaced by a bot.');
	s = s.replace(/^(.+?) เข้าร่วมห้องสำเร็จ$/, '$1 joined the room.');
	s = s.replace(/^เพิ่ม (.+?) เข้าห้องแล้ว$/, 'Added $1 to the room.');
	s = s.replace(/^ลด (.+?) ออกจากห้องแล้ว$/, 'Removed $1 from the room.');

	// 3. ข้อความการจั่ว
	const selfDrawMatch = s.match(/^คุณจั่วได้ (.+?) กับ (.+?) : (ลงได้|ลงไม่ได้)$/);
	if (selfDrawMatch) {
		const v1 = translateSpecialWordsToEn(selfDrawMatch[1]);
		const v2 = translateSpecialWordsToEn(selfDrawMatch[2]);
		const st = selfDrawMatch[3] === 'ลงได้' ? 'Playable' : 'Not playable';
		return `You drew ${v1} and ${v2}: ${st}`;
	}
	s = s.replace(/^(.+?) จั่ว 1 ตัว$/, '$1 drew 1 tile.');
	s = s.replace(/^(.+?) จั่วครบ 3 ตัวแล้ว แต่ยังไม่มีตัวที่วางได้ จบตา$/, '$1 drew 3 tiles with no playable match. Turn ended.');

	// 4. ข้อความ DOMINO! และบทลงโทษ
	s = s.replace(/^(.+?) ประกาศ DOMINO! เหลือ 2 ตัว$/, '$1 declared DOMINO! 2 tiles left.');
	s = s.replace(/^(.+?) ลืมกด DOMINO! ถูกลงโทษ จั่ว 2 ตัว$/, '$1 forgot to press DOMINO! Penalty: draw 2 tiles.');

	// 5. ข้อความการวางหมาก
	const playRegMatch = s.match(/^(.+?) วาง (.+?) กับ (.+?) (เป็นตัวแรก|ที่ปลายซ้าย|ที่ปลายขวา)$/);
	if (playRegMatch) {
		const pName = playRegMatch[1];
		const v1 = translateSpecialWordsToEn(playRegMatch[2]);
		const v2 = translateSpecialWordsToEn(playRegMatch[3]);
		const posMap = {
			'เป็นตัวแรก': 'as the first tile.',
			'ที่ปลายซ้าย': 'on the left end.',
			'ที่ปลายขวา': 'on the right end.'
		};
		return `${pName} placed ${v1} and ${v2} ${posMap[playRegMatch[4]]}`;
	}
	const playSpMatch = s.match(/^(.+?) วาง (.+?) (เป็นตัวแรก|ที่ปลายซ้าย|ที่ปลายขวา)$/);
	if (playSpMatch) {
		const pName = playSpMatch[1];
		const spName = translateSpecialWordsToEn(playSpMatch[2]);
		const posMap = {
			'เป็นตัวแรก': 'as the first tile.',
			'ที่ปลายซ้าย': 'on the left end.',
			'ที่ปลายขวา': 'on the right end.'
		};
		return `${pName} placed ${spName} ${posMap[playSpMatch[3]]}`;
	}

	s = s.replace(/^(.+?) เหลือ 1 ตัว$/, '$1 has 1 tile left.');

	// 6. ข้อความวาง Double ต่อเนื่อง
	const dblMatch = s.match(/^(.+?) วาง Double สามารถเล่นต่อเนื่องได้ที่ฝั่งเดิม(.*)$/);
	if (dblMatch) {
		const pName = dblMatch[1];
		let extra = dblMatch[2] || '';
		extra = extra.replace(/ลงได้ (\d+) ตัว เลือกเลย/, '$1 playable tile(s). Choose one!');
		return `${pName} placed a Double and can chain another tile on the same end! ${extra}`.trim();
	}

	// 7. เอฟเฟกต์พิเศษ
	s = s.replace(/^(.+?) ถูกข้ามตา$/, '$1 is skipped.');
	s = s.replace(/^ทิศทางการเล่นเปลี่ยนเป็น(ตามเข็มนาฬิกา|ทวนเข็มนาฬิกา)$/, (_, d) =>
		`Direction reversed to ${d === 'ตามเข็มนาฬิกา' ? 'clockwise' : 'counter-clockwise'}.`
	);
	s = s.replace(/^(.+?)โดน 2\+ ต้องจั่วเพิ่ม 2 ตัวและถูกข้ามตา$/, '$1 got hit by 2+! Draws 2 tiles and is skipped.');
	s = s.replace(/^(.+?)โดน 3\+ ต้องจั่วเพิ่ม 3 ตัวและถูกข้ามตา$/, '$1 got hit by 3+! Draws 3 tiles and is skipped.');

	// 8. กองจั่วหมด / เกมตัน / จบเกม
	if (s === 'กองจั่วหมดและไม่มีผู้เล่นสามารถดำเนินเกมต่อได้') {
		return 'Draw pile is empty and no players can make a move.';
	}
	s = s.replace(/^กองจั่วหมดแล้ว (.+?) ผ่านรอบ$/, 'Draw pile is empty. $1 passes.');
	if (s === 'เกมตัน! ทำการรวมแต้มเพื่อหาผู้ชนะ') {
		return 'Game is blocked! Tallying remaining points to determine the winner.';
	}
	if (s.startsWith('การแข่งขันจบแล้ว ')) {
		let resStr = s.replace('การแข่งขันจบแล้ว ', 'Match over! ');
		resStr = resStr.replace(/(.+?)เป็นผู้ชนะ เหลือ (\d+) แต้ม\./g, '$1 wins with $2 points left.');
		resStr = resStr.replace(/(.+?)ได้ (\d+) แต้ม แพ้\./g, '$1 has $2 points (Lost).');
		return resStr.trim();
	}

	return s;
}

function getLocalizedMessage(msg) {
	if (!msg) return '';
	return currentLang === 'en' ? translateAnnouncementToEn(msg) : msg;
}

let lastDisplayedStats = null;
let lastDisplayedWinner = '';
let lastRenderedGameState = null;
let lastRenderedRooms = null;

function applyStaticLanguage() {
	const dict = DOMINO_I18N[currentLang] || DOMINO_I18N.th;
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
	localStorage.setItem('domino_lang', currentLang);
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
	const localMsg = getLocalizedMessage(msg);
	announce(localMsg, assertive);
	logEvent(localMsg);
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
	const isEn = currentLang === 'en';
	if(val === 'sleep') return isEn ? 'Sleep' : 'หลับ';
	if(val === 'draw2') return '2+';
	if(val === 'draw3') return '3+';
	if(val === 'reverse') return isEn ? 'Reverse' : 'ย้อนศร';
	return val;
}

function getThaiTileLabel(val) {
	if (val === 'sleep') return 'หลับ';
	if (val === 'draw2') return '2+';
	if (val === 'draw3') return '3+';
	if (val === 'reverse') return 'ย้อนศร';
	return val;
}

// --- Visual FX Layer Functions (Non-blocking) ---
function getPipLayoutHTML(num) {
	if (typeof num !== 'number' || num === 0) return `<div class="pips-grid"></div>`;
	const dots = Array(9).fill(0).map((_, i) => `<div class="pip pip-id-${i}"></div>`);
	let hideIndices = [];
	if(num === 1) hideIndices = [0,1,2,3,5,6,7,8];
	else if(num === 2) hideIndices = [0,1,3,4,5,7,8];
	else if(num === 3) hideIndices = [0,1,3,5,7,8];
	else if(num === 4) hideIndices = [1,3,4,5,7];
	else if(num === 5) hideIndices = [1,3,5,7];
	else if(num === 6) hideIndices = [1,4,7];
	else if(num === 7) hideIndices = [1,7];
	else if(num === 8) hideIndices = [4];
	else if(num === 9) hideIndices = [];
	
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
	fxEl.innerHTML = renderTileHTML(tile, -1, false, false, true).outerHTML;
	
	const startX = sourceRect.left + sourceRect.width/2 - 25;
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
	if(!targetEl) targetEl = document.getElementById('my-cards-container');
	
	if(!deckEl || !targetEl) return;
	const deckRect = deckEl.getBoundingClientRect();
	const targetRect = targetEl.getBoundingClientRect();

	const fxEl = document.createElement('div');
	fxEl.className = 'fx-tile domino-tile';
	fxEl.style.background = 'linear-gradient(145deg, #1f2937, #111827)';
	
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

applyStaticLanguage();
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
		announce(currentLang === 'en'
			? 'Domino room is ready! Wait for friends or add bots now.'
			: 'ห้องโดมิโน่พร้อมแล้ว รอเพื่อนหรือเพิ่มบอทได้ทันที');
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
	if (isHost && currentRoomId) { remove(ref(db, `domino_rooms/${currentRoomId}`)); connections.forEach(c => c.close()); }
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
	document.getElementById('lobby-ready-msg').innerText = currentLang === 'en'
		? 'Up to 6 players per room. Host can press Start Game whenever ready!'
		: 'เล่นได้สูงสุดที่ 6 คนต่อห้อง เจ้าของห้องพร้อมก็กดเริ่มเกมได้นะ เพื่อสนุกกัน';
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
	// ตัวปกติ 0-9 (55 ตัว)
	for (let i = 0; i <= 9; i++) {
		for (let j = i; j <= 9; j++) {
			deck.push({ id: idCount++, left: i, right: j, type: 'regular', name: `${i} กับ ${j}` });
		}
	}
	
	// ตัวพิเศษครบทั้ง 4 แบบ สำหรับเลข 0 ถึง 9 (รวม 40 ตัว)
	const effects = [
		{ effect: 'sleep', name: 'หลับ' },
		{ effect: 'draw2', name: '2+' },
		{ effect: 'draw3', name: '3+' },
		{ effect: 'reverse', name: 'ย้อนศร' }
	];

	for (let num = 0; num <= 9; num++) {
		effects.forEach(ef => {
			deck.push({ 
				id: idCount++, 
				left: num, 
				right: ef.effect, 
				type: ef.effect, 
				name: `${num} กับ ${ef.name}` 
			});
		});
	}

	return deck.sort(() => Math.random() - 0.5);
}

// คะแนนตอนจบเกม: ตัวปกติ = ซ้าย + ขวา | ตัวพิเศษ = 20 + ตัวเลขอีกฝั่ง (left)
function getTileScore(tile) {
	if (tile.type === 'regular') return tile.left + tile.right;
	return 20 + (typeof tile.left === 'number' ? tile.left : 0);
}

// จัดเรียงตัวโดมิโน่ในมือ:
// 1. จัดเป็น 2 กลุ่มหลัก: กลุ่มปกติ (เรียงแต้มน้อย -> มาก) ตามด้วย กลุ่มพิเศษ (เรียง 20+เลข น้อย -> มาก)
// 2. หากเป็นเทิร์นของผู้เล่นคนนั้น (prioritizePlayable = true):
//    ดึงตัวที่ลงได้ขึ้นมาไว้หน้าสุดตามลำดับ:
//    [Double ที่ลงได้] -> [ปกติที่ลงได้] -> [พิเศษที่ลงได้] -> [ปกติที่ลงไม่ได้] -> [พิเศษที่ลงไม่ได้]
function sortPlayerHand(hand, prioritizePlayable = false, lockedSide = null) {
	if (!Array.isArray(hand) || hand.length <= 1) return;

	const getSortPriority = (tile) => {
		const isSpecial = tile.type !== 'regular';
		const isDouble = !isSpecial && tile.left === tile.right;

		if (prioritizePlayable) {
			const playable = isTilePlayable(tile, lockedSide);
			if (playable) {
				if (isDouble) return 0;      // 1. Double ที่ลงได้ (เอาขึ้นหน้าสุดเพื่อทำคอมโบต่อเนื่อง)
				if (!isSpecial) return 1;    // 2. ตัวปกติที่ลงได้ (เรียงแต้มน้อยไปมาก)
				return 2;                    // 3. ตัวพิเศษที่ลงได้ (เรียงแต้มน้อยไปมาก)
			}
		}
		// กลุ่มลงไม่ได้ (หรือนอกเทิร์นของตัวเอง): แยกเป็น 2 กลุ่มตามกติกา
		return isSpecial ? 4 : 3;            // 4. ตัวปกติ (น้อยไปมาก) -> 5. ตัวพิเศษ (น้อยไปมาก)
	};

	hand.sort((a, b) => {
		const prioA = getSortPriority(a);
		const prioB = getSortPriority(b);
		if (prioA !== prioB) return prioA - prioB;

		const scoreA = getTileScore(a);
		const scoreB = getTileScore(b);
		if (scoreA !== scoreB) return scoreA - scoreB;

		// กรณีแต้มเท่ากัน ให้เรียงตามเลขด้านซ้ายและ id เพื่อความเสถียรไม่สลับตำแหน่งไปมา
		const leftA = typeof a.left === 'number' ? a.left : 99;
		const leftB = typeof b.left === 'number' ? b.left : 99;
		if (leftA !== leftB) return leftA - leftB;

		return (a.id || 0) - (b.id || 0);
	});
}

function getTileAria(tile) {
	if (currentLang === 'en') {
		return `${translateTile(tile.left)} and ${translateTile(tile.right)}`;
	}
	return `${tile.name}`;
}

function renderTileHTML(tile, index = -1, playableLeft = false, playableRight = false, forceVisualOnly = false) {
	const isSpecial = tile.type !== 'regular';
	const isDouble = tile.left === tile.right && tile.type === 'regular';
	const btn = document.createElement((index >= 0 && !forceVisualOnly) ? 'button' : 'div');
	const isEn = currentLang === 'en';
	
	let spClass = '';
	if(tile.type === 'sleep') spClass = 'sp-sleep';
	else if(tile.type === 'draw2') spClass = 'sp-draw2';
	else if(tile.type === 'draw3') spClass = 'sp-draw3';
	else if(tile.type === 'reverse') spClass = 'sp-reverse';
	
	btn.className = `domino-tile ${isDouble ? '' : 'horizontal'} ${isSpecial ? 'special-tile ' + spClass : ''} ${(index >= 0 && !playableLeft && !playableRight && !forceVisualOnly) ? 'disabled' : ''}`;
	
	const formatHalf = (val) => {
		if(val === 'sleep') return `<div class="domino-half special-text">${isEn ? 'Sleep' : 'หลับ'}<br>💤</div>`;
		if(val === 'draw2') return `<div class="domino-half special-text">2+<br>⚡</div>`;
		if(val === 'draw3') return `<div class="domino-half special-text">3+<br>💥</div>`;
		if(val === 'reverse') return `<div class="domino-half special-text">${isEn ? 'Reverse' : 'ย้อนศร'}<br>🔄</div>`;
		return `<div class="domino-half">${getPipLayoutHTML(val)}</div>`;
	};

	btn.innerHTML = `${formatHalf(tile.left)}${formatHalf(tile.right)}`;
	
	if (index >= 0 && !forceVisualOnly) {
		const statusStr = isEn
			? (playableLeft || playableRight ? 'Playable' : 'Unplayable')
			: (playableLeft || playableRight ? 'ลงได้' : 'ลงไม่ได้');
		btn.setAttribute('aria-label', `${getTileAria(tile)} ${statusStr}`);
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
	let baseMsg = `ถึงรอบของ ${currentPlayer.name} ขณะนี้เล่น${dirStr} `;
	if (game.leftEnd === null) {
		baseMsg += `กระดานยังว่าง คุณสามารถเริ่มวางได้`;
	} else {
		baseMsg += `ปลายซ้ายคือ ${getThaiTileLabel(game.leftEnd)} ปลายขวาคือ ${getThaiTileLabel(game.rightEnd)}`;
	}
	if (game.doubleChainSide) {
		baseMsg += ` (มีการล็อกฝั่งเล่นต่อที่ปลาย${game.doubleChainSide === 'left' ? 'ซ้าย' : 'ขวา'})`;
	}

	let activePlayerMsg = baseMsg;
	if (!currentPlayer.isBot) {
		const hand = game.playerStates[currentPlayer.id].hand;
		const playableCount = hand.filter(t => isTilePlayable(t, game.doubleChainSide)).length;
		if (playableCount > 0) {
			activePlayerMsg += ` ลงได้ ${playableCount} ตัว เลือกเลย`;
		} else {
			activePlayerMsg += ` ลงไม่ได้จั่วเลย`;
		}
	}

	logEvent(getLocalizedMessage(baseMsg));
	players.forEach(p => {
		if (p.isBot) return;
		const msgToSend = (p.id === currentPlayer.id) ? activePlayerMsg : baseMsg;
		if (p.id === myPeerId) {
			announce(getLocalizedMessage(msgToSend));
		} else if (p.connection && p.connection.open) {
			p.connection.send({ type: 'announce', message: msgToSend, assertive: false });
		}
	});
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

	// แจกเริ่มต้นคนละ 5 ตัว และเรียงมือเป็น 2 กลุ่มทันที
	players.forEach((p, i) => {
		game.playerStates[p.id] = { hand: [], declaredDomino: false };
		for(let k=0; k<5; k++) game.playerStates[p.id].hand.push(game.deck.pop());
		
		game.playerStates[p.id].hand.forEach(t => {
			if(t.type === 'regular' && t.left === t.right && t.left > highestDoubleVal) {
				highestDoubleVal = t.left; starterIdx = i;
			}
		});
		sortPlayerHand(game.playerStates[p.id].hand, false);
	});

	game.turnIndex = starterIdx;
	// จัดเรียงมือของผู้เล่นคนแรกโดยดึงตัวที่ลงได้ขึ้นมาไว้หน้าสุด
	const firstStarter = players[game.turnIndex];
	if (firstStarter && game.playerStates[firstStarter.id]) {
		sortPlayerHand(game.playerStates[firstStarter.id].hand, true, game.doubleChainSide);
	}

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
	else if (data.type === 'announce') {
		const localMsg = getLocalizedMessage(data.message);
		announce(localMsg, data.assertive);
		logEvent(localMsg);
	}
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
	lastRenderedGameState = gState;
	const isEn = currentLang === 'en';
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

	const statusBar = document.getElementById('top-status-bar');
	statusBar.innerHTML = '';
	let ariaStatusBarText = isEn
		? `Direction ${gState.direction === 1 ? 'clockwise' : 'counter-clockwise'}, `
		: `ทิศทาง ${gState.direction === 1 ? 'ตามเข็ม' : 'ทวนเข็ม'}, `;
	const dirDiv = document.createElement('div'); dirDiv.className = 'direction-indicator'; dirDiv.setAttribute('aria-hidden', 'true');
	dirDiv.style.transform = gState.direction === 1 ? 'rotate(0deg)' : 'rotate(180deg)'; dirDiv.textContent = '➡️';
	statusBar.appendChild(dirDiv);

	const mascots = ['🐱', '🐶', '🐰', '🦊', '🐼', '🐸'];
	gState.players.forEach((p, idx) => {
		const pState = gState.playerStates[p.id];
		const isTurn = idx === gState.turnIndex;
		const dispPName = formatPlayerDisplayName(p.name);
		ariaStatusBarText += isEn
			? `${dispPName} has ${pState.cardCount} tiles. `
			: `${dispPName} มี ${pState.cardCount} ตัว. `;
		const charDiv = document.createElement('div');
		charDiv.id = `char-${p.id}`;
		charDiv.className = `character-card ${isTurn ? 'is-turn' : ''}`; charDiv.setAttribute('aria-hidden', 'true');
		const mascot = mascots[idx % 6];
		const countVisualText = isEn ? `${pState.cardCount} left` : `เหลือ ${pState.cardCount} ตัว`;
		const hiddenCountText = isEn ? `Tiles ${pState.cardCount}` : `ไพ่ ${pState.cardCount}`;
		charDiv.innerHTML = `<div style="display: none;"><div class="char-cards-count">${hiddenCountText}</div><div class="char-name">${p.isBot?'🤖 ':'👤 '}${dispPName}</div></div><div class="char-cards-count">${countVisualText}</div><div class="char-name">${mascot} ${dispPName}</div>`;
		statusBar.appendChild(charDiv);
	});
	statusBar.setAttribute('aria-label', ariaStatusBarText);

	document.getElementById('deck-count-visual').textContent = gState.deckCount;
	document.getElementById('deck-aria-label').textContent = isEn
		? `Draw pile has ${gState.deckCount} tiles left`
		: `กองจั่วเหลือ ${gState.deckCount} ตัว`;

	// Render Board Ends visually
	const lCont = document.getElementById('board-left-container');
	const rCont = document.getElementById('board-right-container');
	const lAria = document.getElementById('board-left-aria');
	const rAria = document.getElementById('board-right-aria');
	
	// Cleanup visual nodes but keep ARIA labels
	Array.from(lCont.children).forEach(c => { if(c.className !== 'board-label') c.remove(); });
	Array.from(rCont.children).forEach(c => { if(c.className !== 'board-label') c.remove(); });

	if (gState.leftEnd === null) {
		lAria.textContent = isEn ? 'Board is empty' : 'กระดานว่างเปล่า';
		rAria.textContent = isEn ? 'Board is empty' : 'กระดานว่างเปล่า';
	} else {
		const renderEndVisual = (val) => {
			const d = document.createElement('div');
			d.className = 'domino-end-tile domino-half';
			d.setAttribute('aria-hidden', 'true');
			if (val === 'sleep') {
				d.classList.add('special-tile', 'sp-sleep');
				d.innerHTML = `<div class="special-text">${isEn ? 'Sleep' : 'หลับ'}<br>💤</div>`;
			} else if (val === 'draw2') {
				d.classList.add('special-tile', 'sp-draw2');
				d.innerHTML = `<div class="special-text">2+<br>⚡</div>`;
			} else if (val === 'draw3') {
				d.classList.add('special-tile', 'sp-draw3');
				d.innerHTML = `<div class="special-text">3+<br>💥</div>`;
			} else if (val === 'reverse') {
				d.classList.add('special-tile', 'sp-reverse');
				d.innerHTML = `<div class="special-text">${isEn ? 'Reverse' : 'ย้อนศร'}<br>🔄</div>`;
			} else {
				d.innerHTML = getPipLayoutHTML(val);
			}
			return d;
		};
		lCont.appendChild(renderEndVisual(gState.leftEnd));
		rCont.appendChild(renderEndVisual(gState.rightEnd));
		lAria.textContent = isEn
			? `Left end ${translateTile(gState.leftEnd)}`
			: `ปลายซ้าย ${translateTile(gState.leftEnd)}`;
		rAria.textContent = isEn
			? `Right end ${translateTile(gState.rightEnd)}`
			: `ปลายขวา ${translateTile(gState.rightEnd)}`;
	}

	// Render Hand
	const myContainer = document.getElementById('my-cards-container');
	myContainer.innerHTML = '';
	
	const isMyTurn = (turnPlayer && turnPlayer.id === myPeerId && gState.status === 'playing');
	const activeChainSide = gState.doubleChainSide;
	
	let canPlayAtLeastOne = false;
	gState.myHand.forEach((tile, idx) => {
		let playableL = false, playableR = false;
		if (isMyTurn) {
			if (gState.leftEnd === null) { playableL = true; playableR = true; }
			else {
				if (!activeChainSide || activeChainSide === 'left') {
					if (tile.left === gState.leftEnd || tile.right === gState.leftEnd) playableL = true;
				}
				if (!activeChainSide || activeChainSide === 'right') {
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
		
		// แทรกตัวที่จั่วได้และจัดเรียงใหม่ทันที โดยดึงตัวที่ลงได้ขึ้นมาไว้หน้าสุด
		sortPlayerHand(game.playerStates[peerId].hand, true, game.doubleChainSide);

		broadcastSound('jua');
		fxDrawTile(peerId);
		connections.forEach(c => { if(c.open) c.send({ type: 'fxDraw', peerId: peerId }); });

		const canPlayDrawn = isTilePlayable(tile, game.doubleChainSide);
		const playStatusText = canPlayDrawn ? ' : ลงได้' : ' : ลงไม่ได้';

		players.forEach(p => {
			if (p.id === peerId) {
				const msg = `คุณจั่วได้ ${getThaiTileLabel(tile.left)} กับ ${getThaiTileLabel(tile.right)}${playStatusText}`;
				if (p.connection && p.connection.open) p.connection.send({ type: 'announce', message: msg });
				else if (p.id === myPeerId) announce(getLocalizedMessage(msg));
			} else {
				const msg = `${currentPlayer.name} จั่ว 1 ตัว`;
				if (p.connection && p.connection.open) p.connection.send({ type: 'announce', message: msg });
				else if (p.id === myPeerId) announce(getLocalizedMessage(msg));
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
	const forgotDomino = (state.hand.length === 2 && !state.declaredDomino);

	state.hand.splice(payload.index, 1);

	if (forgotDomino) {
		broadcastAnnounce(`${currentPlayer.name} ลืมกด DOMINO! ถูกลงโทษ จั่ว 2 ตัว`);
		broadcastSound('wl');
		drawTiles(peerId, 2);
		state.declaredDomino = false;
	}

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

	// จัดเรียงมือหลังวาง: หากวาง Double และมีตัวเล่นต่อได้ให้ดึงตัวที่เล่นต่อได้ขึ้นหน้าสุด มิฉะนั้นเรียง 2 กลุ่มปกติ
	const willChainDouble = (isDouble && state.hand.length > 0 && canPlayAny(state.hand, game.doubleChainSide));
	sortPlayerHand(state.hand, willChainDouble, game.doubleChainSide);

	let actionStr = '';
	if (tile.type === 'regular') {
		actionStr = `${currentPlayer.name} วาง ${tile.left} กับ ${tile.right} ${playedSideStr}`;
	} else {
		actionStr = `${currentPlayer.name} วาง ${getThaiTileLabel(tile.type)} ${playedSideStr}`;
	}
	
	broadcastAnnounce(actionStr);
	broadcastSound('select');

	game.consecutivePasses = 0;
	broadcastGameState();

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
			// จัดเรียงอีกครั้งเผื่อมีตัวจั่วลงโทษเพิ่มเข้ามาตอนวาง Double
			sortPlayerHand(st.hand, true, game.doubleChainSide);
			const baseDoubleMsg = `${currentPlayer.name} วาง Double สามารถเล่นต่อเนื่องได้ที่ฝั่งเดิม`;
			let activeDoubleMsg = baseDoubleMsg;
			if (!currentPlayer.isBot) {
				const playableCount = st.hand.filter(t => isTilePlayable(t, game.doubleChainSide)).length;
				activeDoubleMsg += ` ลงได้ ${playableCount} ตัว เลือกเลย`;
			}
			logEvent(getLocalizedMessage(baseDoubleMsg));
			players.forEach(p => {
				if (p.isBot) return;
				const msgToSend = (p.id === currentPlayer.id) ? activeDoubleMsg : baseDoubleMsg;
				if (p.id === myPeerId) {
					announce(getLocalizedMessage(msgToSend));
				} else if (p.connection && p.connection.open) {
					p.connection.send({ type: 'announce', message: msgToSend, assertive: false });
				}
			});
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
	const isCurrentTurn = (players[game.turnIndex] && players[game.turnIndex].id === playerId);
	sortPlayerHand(st.hand, isCurrentTurn, game.doubleChainSide);
}

function isTilePlayable(tile, lockedSide = null) {
	if (game.leftEnd === null) return true;
	if (lockedSide === 'left') {
		return tile.left === game.leftEnd || tile.right === game.leftEnd;
	}
	if (lockedSide === 'right') {
		return tile.left === game.rightEnd || tile.right === game.rightEnd;
	}
	return tile.left === game.leftEnd || tile.right === game.leftEnd || tile.left === game.rightEnd || tile.right === game.rightEnd;
}

function canPlayAny(hand, lockedSide = null) {
	if (game.leftEnd === null) return hand.length > 0;
	return hand.some(t => isTilePlayable(t, lockedSide));
}

function advanceTurn() {
	let steps = 1 + (game.pendingSkips || 0);
	game.pendingSkips = 0;
	game.drawCountThisTurn = 0;
	game.doubleChainSide = null;
	game.turnIndex = (game.turnIndex + (game.direction * steps) + (players.length * 10)) % players.length;
	
	const nextPlayer = players[game.turnIndex];

	// จัดเรียงมือผู้เล่นทั้งหมด:
	// - ผู้เล่นที่กำลังจะถึงตา: ดึงตัวที่ลงได้ขึ้นมาไว้หน้าสุด (Double -> ปกติ -> พิเศษ) แล้วตามด้วยตัวที่ลงไม่ได้
	// - ผู้เล่นคนอื่น: จัดเรียงเป็น 2 กลุ่ม (ปกติ -> พิเศษ) จากน้อยไปมาก
	players.forEach(p => {
		const st = game.playerStates[p.id];
		if (st && st.hand) {
			sortPlayerHand(st.hand, p.id === nextPlayer.id, null);
		}
	});

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
	await new Promise(r => setTimeout(r, 2000));

	while (game.drawCountThisTurn < 3 && !canPlayAny(state.hand, game.doubleChainSide) && game.deck.length > 0) {
		const tile = game.deck.pop();
		state.hand.push(tile);
		state.declaredDomino = false;
		game.drawCountThisTurn++;
		sortPlayerHand(state.hand, true, game.doubleChainSide);
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
	lastDisplayedWinner = winnerName;
	lastDisplayedStats = resultStats;
	switchScreen('screen-result', 'title-result');
	stopBGM();
	const isEn = currentLang === 'en';
	
	let winnerStat = resultStats.find(r => r.isWinner) || resultStats[0];
	let losers = resultStats.filter(r => !r.isWinner);
	const dispWinnerName = formatPlayerDisplayName(winnerName);
	
	let html = isEn
		? `<div style="color: var(--focus-ring); margin-bottom: 10px;">👑 ${dispWinnerName} wins with ${winnerStat.points} points left</div>`
		: `<div style="color: var(--focus-ring); margin-bottom: 10px;">👑 ${dispWinnerName}เป็นผู้ชนะ เหลือ ${winnerStat.points} แต้ม</div>`;
	losers.forEach(l => {
		const dispLoserName = formatPlayerDisplayName(l.name);
		html += isEn
			? `<div style="color: #ffcdd2; margin-bottom: 10px;">❌ ${dispLoserName} has ${l.points} points (Lost)</div>`
			: `<div style="color: #ffcdd2; margin-bottom: 10px;">❌ ${dispLoserName}ได้ ${l.points} แต้ม แพ้</div>`;
	});
	
	document.getElementById('winner-text').innerHTML = html;
	document.getElementById('result-status-bar').innerHTML = '';
}

// --- Keyboard Shortcuts ---
document.addEventListener('keydown', (e) => {
	if (!document.getElementById('screen-game').classList.contains('active')) return;
	if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
	const isEn = currentLang === 'en';

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
		let msg = isEn ? "All players status: " : "สถานะตัวละครทั้งหมด: ";
		const chars = document.querySelectorAll('#top-status-bar .character-card');
		let parts = [];
		chars.forEach(c => {
			const countText = c.querySelector('.char-cards-count').textContent.replace('ไพ่ ', ''); 
			const nameText = c.querySelector('.char-name').textContent.replace(/[🤖👤🐱🐶🐰🦊🐼🐸]/g, '').trim();
			parts.push(isEn ? `${nameText} has ${countText}` : `${nameText}เหลือ ${countText}`);
		});
		announce(msg + parts.join(' '));
	}
});