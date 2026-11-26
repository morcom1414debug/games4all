import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getDatabase, ref, set, onValue, remove, onDisconnect, runTransaction } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-database.js";

// Global Utility & Guard States
window.joinOrder = []; 
window.isStartingRound = false; 
window.hasAnnouncedMyColor = false; 

// =========================================================
// Bilingual Localization System (TH / EN) - Storage Key: cang_lang
// =========================================================
const LANG_STORAGE_KEY = 'cang_lang';

function detectInitialLanguage() {
    try {
        const saved = localStorage.getItem(LANG_STORAGE_KEY);
        if (saved === 'th' || saved === 'en') return saved;
    } catch (e) {}
    const browserLang = (navigator.language || navigator.userLanguage || 'en').toLowerCase();
    return browserLang.startsWith('th') ? 'th' : 'en';
}

let currentLang = detectInitialLanguage();
let originalThaiManualHTML = null;
let lastResultRawData = null;
let roomsSnapshotCache = null;

const EN_SUITS = { '♠': 'Spades', '♥': 'Hearts', '♦': 'Diamonds', '♣': 'Clubs' };
const EN_RANKS = {
    'A': 'Ace', '2': '2', '3': '3', '4': '4', '5': '5', '6': '6',
    '7': '7', '8': '8', '9': '9', '10': '10', 'J': 'Jack', 'Q': 'Queen', 'K': 'King'
};

const SPECIAL_WIN_LABELS = {
    th: {
        '3A': '3A',
        '50': '50',
        'ตอง': 'ตอง',
        'เรียง': 'เรียง',
        'สี': 'สี',
        '3คู่': '3 คู่'
    },
    en: {
        '3A': '3A',
        '50': '50 Pts',
        'ตอง': 'Three of a Kind',
        'เรียง': 'Straight',
        'สี': 'Flush',
        '3คู่': '3 Pairs'
    }
};

function getSpecialWinDisplayName(typeKey) {
    if (!typeKey) return '';
    const dict = SPECIAL_WIN_LABELS[currentLang] || SPECIAL_WIN_LABELS.th;
    return dict[typeKey] || typeKey;
}

const I18N = {
    th: {
        pageTitle: 'แคง1.4',
        homeAria: 'กลับสู่ชุมชน',
        langToggleText: 'EN',
        langToggleAria: 'Switch language to English',
        transitionDefault: 'กำลังเปลี่ยนตาถัดไป...',
        disconnectHeading: 'การเชื่อมต่อกับเจ้าของห้องขาดหาย',
        nameEntryTitle: 'ยินดีต้อนรับสู่ "แคง" ออนไลน์ 🎴',
        nameEntryDesc: 'กรุณาตั้งชื่อของคุณเพื่อเข้าสู่เกม',
        playerNameLabel: 'ชื่อผู้เล่นของคุณ:',
        playerNamePlaceholder: 'พิมพ์ชื่อผู้เล่น...',
        btnConfirmName: 'ยืนยันชื่อผู้เล่น',
        setupTitle: '"แคง" ออนไลน์ 🎴',
        setupDesc: 'จะสร้างหรือเลือกเข้าร่วมห้องเพื่อดวลแต้ม เลือกเลย',
        entertainmentNote: 'สำหรับการสันทนาการเท่านั้น',
        btnCreateRoom: 'สร้างห้องเล่นแคง',
        btnOpenManual: 'คู่มือการเล่น',
        roomListTitle: 'ห้องที่กำลังรอผู้เล่น:',
        noRoomSearching: 'กำลังค้นหาห้อง...',
        noRoomEmpty: 'ไม่มีห้องว่าง (กดสร้างห้องได้เลย)',
        roomItemAria: (rId, count) => `เข้าร่วมห้อง ${rId} มีผู้เล่น ${count} จาก 6 คน`,
        roomItemLabel: (rId, count) => `ห้อง ${rId} (${count}/6 คน) `,
        roomItemJoinArrow: 'เข้า ➔',
        manualTitle: '📖 คู่มือการเล่น ไพ่แคง',
        btnCloseManual: 'ปิดหน้าคู่มือ กลับสู่หน้าหลัก',
        manualOpenedAnnounce: 'เปิดหน้าคู่มือการเล่นแล้ว',
        manualClosedAnnounce: 'ปิดหน้าคู่มือการเล่น กลับสู่หน้าหลัก',
        waitingRoomAria: 'ห้องรอผู้เล่น',
        btnLeaveWaiting: '🚪 ออก',
        btnLeaveWaitingAria: 'ออกจากห้อง กลับหน้าหลัก',
        waitingRoomTitle: '🎴 ห้องรอผู้เล่น',
        roomIdLabel: 'รหัสห้อง (Game ID):',
        botSelectTitle: 'เลือกจำนวนบอทในห้อง:',
        btnBotDecAria: 'ลดจำนวนบอท',
        btnBotIncAria: 'เพิ่มจำนวนบอท',
        botCountUnit: (n) => `${n} ตัว`,
        waitingPlayersTitle: 'รายชื่อผู้เล่นและบอท (เริ่มต้น 200 เหรียญ):',
        btnHostStart: 'แจกไพ่เริ่มเกม',
        guestWaitingText: 'รอเจ้าของห้องแจกไพ่...',
        gameScreenAria: 'โต๊ะเล่นไพ่แคง',
        deckLabel: 'กองจั่ว',
        deckUnit: 'ใบ',
        deckAria: (count) => `กองไพ่สำหรับจั่ว เหลือ ${count} ใบ`,
        welcomeTableAnnounce: 'ยินดีต้อนรับสู่โต๊ะไพ่แคง!',
        myHandAria: 'ไพ่ในมือของคุณ',
        btnDraw: 'จั่วไพ่',
        btnEndTurn: 'จบตา',
        btnKang: 'แคง!',
        btnKang25: 'แคง 25!',
        btnSkipPre: 'ไม่จ้า',
        btnNextRound: 'เริ่มรอบใหม่',
        btnBackHomeReload: 'กลับหน้าหลัก',
        btnResultHome: '🏠 กลับสู่ชุมชน',
        guestWaitingNextRound: 'รอเจ้าของห้องเริ่มรอบใหม่...',
        tempBotHeading: 'คุณถูกเปลี่ยนเป็นบอทชั่วคราว',
        tempBotP1: 'เนื่องจากคุณหมดเวลาเล่นครบ 3 ครั้ง ระบบจึงให้บอทเล่นแทนเพื่อให้เกมดำเนินต่อไป',
        tempBotP2: 'คุณต้องการกลับมาเล่นต่อหรือไม่?',
        tempBotCountdown: (sec) => `กรุณาเลือกภายใน ${sec} วินาที มิฉะนั้นระบบจะถือว่าคุณเลือก "ไม่ใช่"`,
        btnBotYes: 'ใช่',
        btnBotNo: 'ไม่ใช่',
        errEmptyName: 'กรุณากรอกชื่อผู้เล่นก่อนเข้าสู่เกม',
        welcomeLobbyAnnounce: (name) => `ยินดีต้อนรับ ${name} เข้าสู่หน้าเลือกห้องเล่นแคง`,
        creatingRoomAnnounce: 'กำลังสร้างห้อง โปรดรอสักครู่',
        roomCreatedAnnounce: (rId) => `สร้างห้องสำเร็จ รหัสห้องคือ ${rId}`,
        joiningRoomAnnounce: (rId) => `กำลังเข้าร่วมห้อง ${rId}`,
        joinedRoomAnnounce: 'เข้าร่วมห้องสำเร็จ รอแจกไพ่',
        preGameTurnInd: 'เตรียมตัว! เช็คไพ่ในมือ (40 วินาที)',
        yourTurnInd: (name) => `🔥 ถึงรอบของ ${name} แล้ว!`,
        otherTurnInd: (name) => `ถึงรอบของ ${name} แล้ว`,
        topCardEmpty: 'ยังไม่มีไพ่บนกอง',
        topCardEmptyFull: 'ยังไม่มีไพ่บนกองทิ้ง',
        topCardCurrent: (cardName) => `ไพ่บนกองทิ้งปัจจุบันคือ ${cardName}`,
        playHintCan: (count) => `ลงได้ ${count} ใบ`,
        playHintCant: 'ลงไม่ได้จั่วเลย',
        yourTurnAnnounce: (name, topTxt, hint) => `ถึงรอบของ ${name} แล้ว กรุณาลงมือเล่น (ไพ่บนกองคือ ${topTxt}) ${hint}`,
        otherTurnAnnounce: (name, topTxt) => `ถึงรอบของ ${name} แล้ว (ไพ่บนกองคือ ${topTxt})`,
        statusPlayerAria: (name, pts) => `${name} มี ${pts} เหรียญ. `,
        statusCoinsText: (pts) => `${pts} เหรียญ`,
        statusGroupAria: (list) => `สถานะผู้เล่นทั้งหมด: ${list}`,
        discardPileTopAria: (cardName) => `ไพ่กองทิ้งใบบนสุดคือ ${cardName}`,
        discardPileEmptyHtml: 'กองทิ้งว่างเปล่า',
        discardPileEmptyAria: 'กองไพ่ทิ้งยังว่างเปล่า',
        hintFlow: ' กดเพื่อไหลไพ่ใบนี้',
        hintDiscard: ' กดเพื่อทิ้งไพ่ใบนี้',
        hintDiscardMore: ' กดเพื่อทิ้งเพิ่มเติม (เลขเดียวกัน)',
        cardAria: (cardName, hint) => `ไพ่ ${cardName}${hint}`,
        preGameSummary: (base, sum, specialKey) => `${base} ผลรวม ${sum} แต้ม${specialKey ? ` คุณใช้ ${getSpecialWinDisplayName(specialKey)} ได้` : ''}`,
        kangAnim1_cang: (name) => `โอ้ ${name} เอ่ยคำว่า แคง`,
        kangAnim1_cang25: (name) => `โอ้ ${name} เอ่ยคำว่า แคง25`,
        kangAnim1_knock: () => 'KNOCK!',
        kangAnim1_special: () => 'SPECIAL WIN!',
        kangAnim2_cang: () => 'เปิดไพ่ทุกคนเพื่อวัดดวงกัน',
        kangAnim2_knock: (name) => `น็อกแล้วโดย ${name}`,
        kangAnim2_special: (name) => `กติกาพิเศษโดย ${name}`
    },
    en: {
        pageTitle: 'Kang 1.4',
        homeAria: 'Back to Community',
        langToggleText: 'TH',
        langToggleAria: 'เปลี่ยนภาษาเป็นภาษาไทย',
        transitionDefault: 'Next turn...',
        disconnectHeading: 'Connection to room host lost',
        nameEntryTitle: 'Welcome to "Kang" Online 🎴',
        nameEntryDesc: 'Please enter your name to join the game',
        playerNameLabel: 'Your Player Name:',
        playerNamePlaceholder: 'Enter player name...',
        btnConfirmName: 'Confirm Name',
        setupTitle: '"Kang" Online 🎴',
        setupDesc: 'Create a room or join one to start playing',
        entertainmentNote: 'For entertainment purposes only',
        btnCreateRoom: 'Create Kang Room',
        btnOpenManual: 'How to Play',
        roomListTitle: 'Available Rooms:',
        noRoomSearching: 'Searching for rooms...',
        noRoomEmpty: 'No rooms available (Create one now!)',
        roomItemAria: (rId, count) => `Join room ${rId} with ${count} of 6 players`,
        roomItemLabel: (rId, count) => `Room ${rId} (${count}/6) `,
        roomItemJoinArrow: 'Join ➔',
        manualTitle: '📖 Kang Cards — How to Play',
        btnCloseManual: 'Close Guide & Return',
        manualOpenedAnnounce: 'Opened How to Play guide',
        manualClosedAnnounce: 'Closed How to Play guide, returned to main menu',
        waitingRoomAria: 'Waiting Room',
        btnLeaveWaiting: '🚪 Leave',
        btnLeaveWaitingAria: 'Leave room and return to main menu',
        waitingRoomTitle: '🎴 Waiting Room',
        roomIdLabel: 'Room Code (Game ID):',
        botSelectTitle: 'Select Number of Bots:',
        btnBotDecAria: 'Decrease bot count',
        btnBotIncAria: 'Increase bot count',
        botCountUnit: (n) => `${n} ${n === 1 ? 'Bot' : 'Bots'}`,
        waitingPlayersTitle: 'Players & Bots (Starting 200 Coins):',
        btnHostStart: 'Deal & Start Game',
        guestWaitingText: 'Waiting for host to deal cards...',
        gameScreenAria: 'Kang Card Table',
        deckLabel: 'Draw Deck',
        deckUnit: 'Cards',
        deckAria: (count) => `Draw deck with ${count} cards remaining`,
        welcomeTableAnnounce: 'Welcome to the Kang table!',
        myHandAria: 'Your Hand',
        btnDraw: 'Draw',
        btnEndTurn: 'End Turn',
        btnKang: 'Kang!',
        btnKang25: 'Kang 25!',
        btnSkipPre: 'Pass / Ready',
        btnNextRound: 'Next Round',
        btnBackHomeReload: 'Back to Main Menu',
        btnResultHome: '🏠 Back to Community',
        guestWaitingNextRound: 'Waiting for host to start next round...',
        tempBotHeading: 'Switched to Temporary Bot',
        tempBotP1: 'You timed out 3 times, so a bot is temporarily playing for you to keep the game moving.',
        tempBotP2: 'Would you like to resume playing?',
        tempBotCountdown: (sec) => `Please choose within ${sec} seconds, otherwise "No" will be selected automatically.`,
        btnBotYes: 'Yes',
        btnBotNo: 'No',
        errEmptyName: 'Please enter your player name before continuing',
        welcomeLobbyAnnounce: (name) => `Welcome ${name} to the Kang room lobby`,
        creatingRoomAnnounce: 'Creating room, please wait...',
        roomCreatedAnnounce: (rId) => `Room created! Room ID is ${rId}`,
        joiningRoomAnnounce: (rId) => `Joining room ${rId}...`,
        joinedRoomAnnounce: 'Joined room! Waiting for cards to be dealt',
        preGameTurnInd: 'Get ready! Check your hand (40s)',
        yourTurnInd: (name) => `🔥 Your turn, ${name}!`,
        otherTurnInd: (name) => `${name}'s turn`,
        topCardEmpty: 'No card on discard pile',
        topCardEmptyFull: 'Discard pile is currently empty',
        topCardCurrent: (cardName) => `Current top discard card is ${cardName}`,
        playHintCan: (count) => `${count} playable`,
        playHintCant: 'No match, draw a card',
        yourTurnAnnounce: (name, topTxt, hint) => `Your turn, ${name}. Make your move (Top card: ${topTxt}). ${hint}`,
        otherTurnAnnounce: (name, topTxt) => `${name}'s turn (Top card: ${topTxt})`,
        statusPlayerAria: (name, pts) => `${name} has ${pts} coins. `,
        statusCoinsText: (pts) => `${pts} Coins`,
        statusGroupAria: (list) => `All players status: ${list}`,
        discardPileTopAria: (cardName) => `Top discard card is ${cardName}`,
        discardPileEmptyHtml: 'Discard Empty',
        discardPileEmptyAria: 'Discard pile is empty',
        hintFlow: ' — Press to Match this card',
        hintDiscard: ' — Press to discard this card',
        hintDiscardMore: ' — Press to discard extra card (same rank)',
        cardAria: (cardName, hint) => `${cardName}${hint}`,
        preGameSummary: (base, sum, specialKey) => `${base} Total: ${sum} pts.${specialKey ? ` You can use ${getSpecialWinDisplayName(specialKey)}!` : ''}`,
        kangAnim1_cang: (name) => `Whoa! ${name} calls Kang!`,
        kangAnim1_cang25: (name) => `Whoa! ${name} calls Kang 25!`,
        kangAnim1_knock: () => 'KNOCK!',
        kangAnim1_special: () => 'SPECIAL WIN!',
        kangAnim2_cang: () => 'Reveal all hands to see who wins!',
        kangAnim2_knock: (name) => `Knock by ${name}!`,
        kangAnim2_special: (name) => `Special Win by ${name}!`
    }
};

const EN_MANUAL_HTML = `
    <h3>Objective</h3>
    <p>Reduce the total points in your hand as low as possible, then call <b>"Kang!"</b> to win and collect coins from every other player.</p>

    <h3>Rules &amp; Gameplay</h3>
    <ul>
        <li><b>Starting Coins (200 Coins):</b> Every player starts with 200 coins and is dealt 5 cards.</li>
        <li><b>Pre-Game Preview Phase (40 Seconds):</b>
            <ul>
                <li>The first player of the round may press <b>"Kang!"</b> to immediately reveal all hands, or press <b>"Kang 25!"</b> if holding 25 or more points.</li>
                <li>Other players holding 25 or more points may press <b>"Kang 25!"</b> to challenge for double stakes.</li>
                <li>If 40 seconds pass with no call (or all players press <b>"Pass / Ready"</b>), the round begins and the first player draws a card.</li>
            </ul>
        </li>
        <li><b>Normal Turn:</b> On your turn, either <b>"Draw"</b> 1 card from the deck or <b>"Match"</b> a card of the same rank as the top discard card. If you drew, discard 1 card (plus any additional cards of the exact same rank) and press <b>"End Turn"</b>.</li>
        <li><b>Matching (Match):</b> If you hold a card matching the rank of the latest discarded card, you can play it directly without drawing and immediately collect 10 coins (20 coins for an Ace) from the player who discarded that card.</li>
    </ul>

    <h3>Card Values</h3>
    <ul>
        <li><b>A (Ace):</b> 1 point</li>
        <li><b>2 to 10:</b> Face value</li>
        <li><b>J, Q, K:</b> 10 points each</li>
    </ul>

    <h3>Scoring &amp; Coin Payouts</h3>
    <ul>
        <li><b>Kang! Succeeded:</b> If you have the lowest score (or tied for lowest), you win! Collect base coins from every player (10 coins + 10 bonus coins per Ace in your hand).</li>
        <li><b>Kang 25! Succeeded:</b> Win double the base payout from every player.</li>
        <li><b>Busted Kang!:</b> If you call Kang! or Kang 25! and another player has a lower score, you pay the full penalty to the lowest-scoring player.</li>
        <li><b>Knock! (Empty Hand):</b> Knocking after drawing awards 20 coins from each player; Knocking by Matching awards 10 coins from each player.</li>
    </ul>

    <h3>Shortcut Keys</h3>
    <ul>
        <li><b>Alt+P</b> Draw a card</li>
        <li><b>Alt+E</b> End turn</li>
        <li><b>Alt+K</b> Call Kang!</li>
        <li><b>Alt+L</b> Call Kang 25!</li>
        <li><b>Alt+C</b> Announce top discard card</li>
    </ul>

    <h3>Special Wins</h3>
    <ul>
        <li><b>3A:</b> Hold three Aces (A) to claim an instant Special Win at any point in the game.</li>
        <li><b>50 Pts:</b> Hold cards totaling exactly 50 points to claim an instant Special Win.</li>
        <li><b>Three of a Kind:</b> Hold 3 or 4 cards of the same rank to claim an instant Special Win.</li>
        <li><b>Straight:</b> Hold 5 consecutive cards in rank order to claim an instant Special Win.</li>
        <li><b>Flush:</b> Hold 5 cards of the same suit to claim an instant Special Win.</li>
        <li><b>3 Pairs:</b> Hold 3 pairs (6 cards total after your first draw) to claim an instant Special Win.</li>
    </ul>

    <p><b>Note:</b> When a Special Win condition is met, a golden Special button appears so you can claim victory immediately. All Special Wins except <b>3A</b> are valid only during each player's first turn.</p>
`;

function t(key, ...args) {
    const dict = I18N[currentLang] || I18N.th;
    const val = dict[key] !== undefined ? dict[key] : I18N.th[key];
    return typeof val === 'function' ? val(...args) : val;
}

function formatCardName(rank, suit) {
    if (currentLang === 'en') {
        return `${EN_RANKS[rank] || rank} of ${EN_SUITS[suit] || suit}`;
    }
    return `${THAI_RANKS[rank] || rank} ${THAI_SUITS[suit] || suit}`;
}

function translateThaiCardPhrase(thaiCardPhrase) {
    if (!thaiCardPhrase) return '';
    const trimmed = thaiCardPhrase.trim();
    for (const r of Object.keys(THAI_RANKS)) {
        for (const s of Object.keys(THAI_SUITS)) {
            const thStr = `${THAI_RANKS[r]} ${THAI_SUITS[s]}`;
            if (trimmed === thStr) {
                return `${EN_RANKS[r]} of ${EN_SUITS[s]}`;
            }
        }
    }
    return trimmed;
}

function localizeDynamicMessage(rawMsg) {
    if (typeof rawMsg !== 'string') return '';
    const withNames = rawMsg.replace(/\[PID:(.*?)\]/g, (match, id) => resolveName(id));
    if (currentLang !== 'en') return withNames;

    let m;
    if (withNames === 'เปิดรอบใหม่! มีใครจะแคง 25 ไหม?') {
        return 'New round! Anyone calling Kang 25?';
    }
    if (withNames === 'ไหลไพ่สำเร็จ! รอสักครู่...') {
        return 'Match successful! Please wait...';
    }
    if (withNames === 'กำลังเปลี่ยนตาถัดไป...') {
        return 'Next turn...';
    }
    if (withNames === 'จบเกมด้วยกติกาพิเศษ! กำลังสรุปผล...') {
        return 'Special Win! Calculating results...';
    }
    if (withNames === 'จบเกม! กำลังสรุปผล...') {
        return 'Round over! Calculating results...';
    }

    m = withNames.match(/^📢 หมดเวลาดูไพ่ 40 วินาที ไม่มีใครกดแคง! ขอเชิญ (.+?) จั่วไพ่ใบแรกและทิ้งลงกองเพื่อเริ่มเกมได้เลยครับ$/);
    if (m) return `📢 40s preview ended — no one called Kang! ${m[1]} goes first: draw a card and discard to start.`;

    m = withNames.match(/^📢 ผู้เล่นทุกคนพร้อมแล้ว! ขอเชิญ (.+?) จั่วไพ่ใบแรกและทิ้งลงกองเพื่อเริ่มเกมได้เลยครับ$/);
    if (m) return `📢 Everyone is ready! ${m[1]} goes first: draw a card and discard to start.`;

    m = withNames.match(/^(.+?) เข้าร่วมห้องใหม่ รวม (\d+) คน$/);
    if (m) return `${m[1]} joined the room (${m[2]} players total)`;

    m = withNames.match(/^(.+?) ออกจากห้องรอผู้เล่น$/);
    if (m) return `${m[1]} left the waiting room`;

    m = withNames.match(/^(.+?) ออกจากห้อง ระบบเปลี่ยนเป็นบอทเล่นแทนจนจบเกม$/);
    if (m) return `${m[1]} left the room. Replaced by a bot for the rest of the game.`;

    m = withNames.match(/^✅ (.+?) กลับมาควบคุมการเล่นตามปกติแล้ว$/);
    if (m) return `✅ ${m[1]} has resumed manual control.`;

    m = withNames.match(/^⏰ (.+?) หมดเวลาเล่น \(ครั้งที่ 1\/3\) ระบบเล่นแทนอัตโนมัติ$/);
    if (m) return `⏰ ${m[1]} timed out (1/3). Auto-playing turn.`;

    m = withNames.match(/^⏰ (.+?) หมดเวลาเล่น 2 ครั้งติดต่อกัน! ตาถัดไปจะลดเวลาเหลือ 20 วินาที$/);
    if (m) return `⏰ ${m[1]} timed out twice in a row! Next turn reduced to 20s.`;

    m = withNames.match(/^🤖 (.+?) หมดเวลาครบ 3 ครั้ง ระบบเปลี่ยนเป็นบอทชั่วคราวเพื่อให้เกมดำเนินต่อไป$/);
    if (m) return `🤖 ${m[1]} timed out 3 times. Switched to a temporary bot.`;

    m = withNames.match(/^(.+?) ไหลไพ่ (.+?)! ได้รับ (\d+) เหรียญ จาก (.+)$/);
    if (m) return `${m[1]} matched ${translateThaiCardPhrase(m[2])}! Won ${m[3]} coins from ${m[4]}`;

    m = withNames.match(/^(.+?) จั่วไพ่ 1 ใบ$/);
    if (m) return `${m[1]} drew 1 card`;

    m = withNames.match(/^(.+?) จั่วได้ไพ่ (.+)$/);
    if (m) return `${m[1]} drew ${translateThaiCardPhrase(m[2])}`;

    m = withNames.match(/^(.+?) วางไพ่ (.+?) เพิ่มเติม \(เลขเดียวกัน\)$/);
    if (m) return `${m[1]} discarded an extra ${translateThaiCardPhrase(m[2])} (same rank)`;

    m = withNames.match(/^(.+?) วางไพ่ (.+)$/);
    if (m) return `${m[1]} discarded ${translateThaiCardPhrase(m[2])}`;

    m = withNames.match(/^(.+?) กดยืนยันจบตา$/);
    if (m) return `${m[1]} ended their turn`;

    m = withNames.match(/^(.+?) ชนะพิเศษกติกา: (.+?)!$/);
    if (m) return `${m[1]} Special Win: ${getSpecialWinDisplayName(m[2])}!`;

    m = withNames.match(/^(.+?) เปิดไพ่ชนะกติกา (.+)$/);
    if (m) return `${m[1]} revealed a ${getSpecialWinDisplayName(m[2])} hand`;

    m = withNames.match(/^(.+?) จั่วแล้วทิ้งไพ่หมดมือ! \(ชนะรับ (\d+) เหรียญ จากทุกคน\)$/);
    if (m) return `${m[1]} Knocked after drawing! (Won ${m[2]} coins from each player)`;

    m = withNames.match(/^(.+?) ไหลไพ่หมดมือน็อก! \(ชนะรับ (\d+) เหรียญ จากทุกคน\)$/);
    if (m) return `${m[1]} Knocked by Matching! (Won ${m[2]} coins from each player)`;

    m = withNames.match(/^(.+?) น็อค!! \(ชนะรับ (\d+) เหรียญ จากทุกคน\)$/);
    if (m) return `${m[1]} Knocked! (Won ${m[2]} coins from each player)`;

    m = withNames.match(/^กองไพ่หมด! (.+?) แต้มต่ำสุด (\d+) แต้ม \(ชนะ\)$/);
    if (m) return `Deck empty! ${m[1]} has the lowest score (${m[2]} pts) and wins!`;

    m = withNames.match(/^(.+?) แคง 25 สำเร็จ!$/);
    if (m) return `${m[1]} Kang 25! Succeeded!`;

    m = withNames.match(/^(.+?) แคง สำเร็จ!$/);
    if (m) return `${m[1]} Kang! Succeeded!`;

    m = withNames.match(/^(.+?) แคง 25 แหก! \((.+?) ชนะ\)$/);
    if (m) return `${m[1]} Busted Kang 25! (${m[2]} wins)`;

    m = withNames.match(/^(.+?) แคง แหก! \((.+?) ชนะ\)$/);
    if (m) return `${m[1]} Busted Kang! (${m[2]} wins)`;

    m = withNames.match(/^(.+?) เปิดไพ่มาแล้วได้แต้มรวม (\d+) แต้ม$/);
    if (m) return `${m[1]} revealed ${m[2]} pts`;

    m = withNames.match(/^(.+?) ได้รับ (\d+) เหรียญ$/);
    if (m) return `${m[1]} won ${m[2]} coins`;

    m = withNames.match(/^(.+?) เสีย (\d+) เหรียญ$/);
    if (m) return `${m[1]} lost ${m[2]} coins`;

    m = withNames.match(/^(.+?) ไม่ได้ไม่เสียเหรียญ$/);
    if (m) return `${m[1]} broke even (0 coins)`;

    return withNames;
}

function renderRoomListFromCache() {
    const listContainer = document.getElementById('room-list-container');
    if (!listContainer) return;
    listContainer.innerHTML = '';
    if (roomsSnapshotCache && typeof roomsSnapshotCache === 'object') {
        let hasRoom = false;
        for (let rId in roomsSnapshotCache) {
            let info = roomsSnapshotCache[rId];
            if (info && info.state === 'WAITING' && info.playerCount < 6) {
                hasRoom = true;
                let btn = document.createElement('button');
                btn.className = 'room-item-btn';
                btn.setAttribute('aria-label', t('roomItemAria', rId, info.playerCount));
                btn.innerHTML = `${t('roomItemLabel', rId, info.playerCount)}<span aria-hidden="true">${t('roomItemJoinArrow')}</span>`;
                btn.onclick = () => window.joinRoomDirect(rId);
                listContainer.appendChild(btn);
            }
        }
        if (!hasRoom) {
            listContainer.innerHTML = `<div class="no-room-text" role="status">${t('noRoomEmpty')}</div>`;
        }
    } else {
        listContainer.innerHTML = `<div class="no-room-text" role="status">${isSystemInitialized ? t('noRoomEmpty') : t('noRoomSearching')}</div>`;
    }
}

function applyLanguage(shouldAnnounce = false) {
    document.documentElement.lang = currentLang;
    document.title = t('pageTitle');

    const homeBtn = document.getElementById('btn-home-top');
    if (homeBtn) homeBtn.setAttribute('aria-label', t('homeAria'));

    const langBtn = document.getElementById('btn-lang-toggle');
    if (langBtn) {
        langBtn.innerText = t('langToggleText');
        langBtn.setAttribute('aria-label', t('langToggleAria'));
    }

    const transText = document.getElementById('transition-text');
    if (transText) transText.innerText = t('transitionDefault');

    const dcHeading = document.getElementById('disconnect-heading');
    if (dcHeading) dcHeading.innerText = t('disconnectHeading');

    const dcHomeBtn = document.getElementById('btn-disconnect-home');
    if (dcHomeBtn) dcHomeBtn.setAttribute('aria-label', t('homeAria'));

    const nameTitle = document.getElementById('name-entry-title');
    if (nameTitle) nameTitle.innerText = t('nameEntryTitle');

    const nameDesc = document.getElementById('name-entry-desc');
    if (nameDesc) nameDesc.innerText = t('nameEntryDesc');

    const nameLabel = document.getElementById('player-name-label');
    if (nameLabel) nameLabel.innerText = t('playerNameLabel');

    const nameInput = document.getElementById('player-name-input');
    if (nameInput) nameInput.setAttribute('placeholder', t('playerNamePlaceholder'));

    const confirmBtnSpan = document.querySelector('#btn-confirm-name span');
    if (confirmBtnSpan) confirmBtnSpan.innerText = t('btnConfirmName');

    const setupTitle = document.getElementById('setup-title');
    if (setupTitle) setupTitle.innerText = t('setupTitle');

    const setupDesc = document.getElementById('setup-desc');
    if (setupDesc) setupDesc.innerText = t('setupDesc');

    const entNote = document.getElementById('entertainment-note');
    if (entNote) entNote.innerText = t('entertainmentNote');

    const createRoomSpan = document.querySelector('#btn-create-room span');
    if (createRoomSpan) createRoomSpan.innerText = t('btnCreateRoom');

    const openManualSpan = document.querySelector('#btn-open-manual span');
    if (openManualSpan) openManualSpan.innerText = t('btnOpenManual');

    const roomListTitle = document.getElementById('room-list-title');
    if (roomListTitle) roomListTitle.innerText = t('roomListTitle');

    const noRoomText = document.getElementById('no-room-text');
    if (noRoomText) noRoomText.innerText = t('noRoomSearching');

    const manualTitle = document.getElementById('manual-title');
    if (manualTitle) manualTitle.innerText = t('manualTitle');

    const manualBody = document.getElementById('manual-content-body');
    if (manualBody) {
        if (originalThaiManualHTML === null) {
            originalThaiManualHTML = manualBody.innerHTML;
        }
        manualBody.innerHTML = (currentLang === 'en') ? EN_MANUAL_HTML : originalThaiManualHTML;
    }

    const closeManualSpan = document.querySelector('#btn-close-manual span');
    if (closeManualSpan) closeManualSpan.innerText = t('btnCloseManual');

    const waitingScreen = document.getElementById('waiting-room-screen');
    if (waitingScreen) waitingScreen.setAttribute('aria-label', t('waitingRoomAria'));

    const leaveWaitingBtn = document.getElementById('btn-leave-waiting');
    if (leaveWaitingBtn) {
        leaveWaitingBtn.innerText = t('btnLeaveWaiting');
        leaveWaitingBtn.setAttribute('aria-label', t('btnLeaveWaitingAria'));
    }

    const waitingTitle = document.getElementById('waiting-room-title');
    if (waitingTitle) waitingTitle.innerText = t('waitingRoomTitle');

    const roomIdLabel = document.getElementById('room-id-label');
    if (roomIdLabel) roomIdLabel.innerText = t('roomIdLabel');

    const botSelectTitle = document.getElementById('bot-select-title');
    if (botSelectTitle) botSelectTitle.innerText = t('botSelectTitle');

    const botDecBtn = document.getElementById('btn-bot-dec');
    if (botDecBtn) botDecBtn.setAttribute('aria-label', t('btnBotDecAria'));

    const botIncBtn = document.getElementById('btn-bot-inc');
    if (botIncBtn) botIncBtn.setAttribute('aria-label', t('btnBotIncAria'));

    const botCountText = document.getElementById('bot-count-text');
    if (botCountText) botCountText.innerText = t('botCountUnit', currentBotCount);

    const waitingPlayersTitle = document.getElementById('waiting-players-title');
    if (waitingPlayersTitle) waitingPlayersTitle.innerText = t('waitingPlayersTitle');

    const hostStartBtn = document.getElementById('btn-host-start');
    if (hostStartBtn) hostStartBtn.innerText = t('btnHostStart');

    const guestWaitTxt = document.getElementById('guest-waiting-text');
    if (guestWaitTxt) guestWaitTxt.innerText = t('guestWaitingText');

    const gameScreen = document.getElementById('game-screen');
    if (gameScreen) gameScreen.setAttribute('aria-label', t('gameScreenAria'));

    const deckLabelText = document.getElementById('deck-label-text');
    if (deckLabelText) deckLabelText.innerText = t('deckLabel');

    const deckUnitText = document.getElementById('deck-unit-text');
    if (deckUnitText) deckUnitText.innerText = t('deckUnit');

    const visAnnouncer = document.getElementById('visible-game-announcer');
    if (visAnnouncer && (visAnnouncer.innerText === 'ยินดีต้อนรับสู่โต๊ะไพ่แคง!' || visAnnouncer.innerText === 'Welcome to the Kang table!')) {
        visAnnouncer.innerText = t('welcomeTableAnnounce');
    }

    const handUi = document.getElementById('my-hand-ui');
    if (handUi) handUi.setAttribute('aria-label', t('myHandAria'));

    const btnDraw = document.getElementById('btn-draw');
    if (btnDraw) btnDraw.innerText = t('btnDraw');

    const btnEnd = document.getElementById('btn-end-turn');
    if (btnEnd) btnEnd.innerText = t('btnEndTurn');

    const btnKang = document.getElementById('btn-kang');
    if (btnKang) btnKang.innerText = t('btnKang');

    const btnKang25 = document.getElementById('btn-kang25');
    if (btnKang25) btnKang25.innerText = t('btnKang25');

    const btnSkip = document.getElementById('btn-skip-pre');
    if (btnSkip) btnSkip.innerText = t('btnSkipPre');

    const btnResultHome = document.getElementById('btn-result-home');
    if (btnResultHome) {
        btnResultHome.innerText = t('btnResultHome');
        btnResultHome.setAttribute('aria-label', t('homeAria'));
    }

    const guestWaitNext = document.getElementById('guest-waiting-next-round');
    if (guestWaitNext) guestWaitNext.innerText = t('guestWaitingNextRound');

    const tempBotHeading = document.getElementById('temporary-bot-heading');
    if (tempBotHeading) tempBotHeading.innerText = t('tempBotHeading');

    const tempBotP1 = document.getElementById('temporary-bot-p1');
    if (tempBotP1) tempBotP1.innerText = t('tempBotP1');

    const tempBotP2 = document.getElementById('temporary-bot-p2');
    if (tempBotP2) tempBotP2.innerText = t('tempBotP2');

    const btnBotYes = document.getElementById('btn-bot-return-yes');
    if (btnBotYes) btnBotYes.innerText = t('btnBotYes');

    const btnBotNo = document.getElementById('btn-bot-return-no');
    if (btnBotNo) btnBotNo.innerText = t('btnBotNo');

    if (roomsSnapshotCache !== null) {
        renderRoomListFromCache();
    }
    if (roomId) {
        updateLobbyUI();
    }
    if (window.currentPublicState) {
        renderClientGame(window.currentPublicState, localPlayerState);
    }
    if (lastResultRawData && document.getElementById('result-modal')?.style.display === 'flex') {
        window.showResultModal(
            lastResultRawData.title,
            lastResultRawData.detail,
            lastResultRawData.isEnd,
            lastResultRawData.winners,
            lastResultRawData.losers,
            true
        );
    }

    if (shouldAnnounce) {
        announce(currentLang === 'en' ? 'Language switched to English' : 'เปลี่ยนภาษาเป็นภาษาไทยแล้ว', false);
    }
}

window.toggleLanguage = () => {
    currentLang = (currentLang === 'th') ? 'en' : 'th';
    try {
        localStorage.setItem(LANG_STORAGE_KEY, currentLang);
    } catch (e) {}
    playSound('select');
    applyLanguage(true);
};

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

const delay = ms => new Promise(res => setTimeout(res, ms));

// =========================================================
// ระบบเรียงไพ่ในมือ (Sort Hand) - ปลอดภัย 100% เพราะเรียงบน Array hand โดยตรง
// =========================================================
const SORT_RANK_ORDER = ['A', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K'];
const SORT_SUIT_ORDER = ['♠', '♥', '♦', '♣'];

function sortPlayerHand(hand, topDiscard = null, hasDrawn = false, hasDiscarded = false, hasFlowedThisTurn = false) {
    if (!Array.isArray(hand) || hand.length <= 1) return hand;
    const shouldPrioritizeFlowable = Boolean(
        topDiscard && topDiscard.rank && !hasDrawn && !hasDiscarded && !hasFlowedThisTurn
    );

    hand.sort((a, b) => {
        if (shouldPrioritizeFlowable) {
            const aMatch = a.rank === topDiscard.rank ? 0 : 1;
            const bMatch = b.rank === topDiscard.rank ? 0 : 1;
            if (aMatch !== bMatch) return aMatch - bMatch;
        }
        const rankDiff = SORT_RANK_ORDER.indexOf(a.rank) - SORT_RANK_ORDER.indexOf(b.rank);
        if (rankDiff !== 0) return rankDiff;
        return SORT_SUIT_ORDER.indexOf(a.suit) - SORT_SUIT_ORDER.indexOf(b.suit);
    });

    return hand;
}

// =========================================================
// Audio Setup (คงไว้ 100% เพิ่มไฟล์เสียงใหม่ และระบบจัดการคิว/AudioContext ส่วนกลาง)
// =========================================================
let audioCtx = null;
let audioResumePromise = null;
let isAudioUnlocked = false;
const rawAudioBuffers = {};
const audioBuffers = {};
const soundLoadStatus = {}; // 'idle' | 'loading' | 'loaded' | 'failed'
const soundLoadPromises = {};
const soundRetryCounts = {};
const MAX_SOUND_RETRY = 3;
let bgmSource = null;
let isBgmWanted = false;

const pendingSoundQueue = [];
let isProcessingSoundQueue = false;
let soundEventSeq = 0;
let lastHandledRemoteSoundId = null;
const recentPlayedSignatureMap = new Map();
const pendingPeerSoundPayloads = new Map();

const soundFiles = {
    '1': 'audio/1.mp3', 'select': 'audio/select.mp3', 'start': 'audio/start.mp3',
    'bgm': 'audio/bgm.mp3', 'jua': 'audio/jua.mp3', 'turn': 'audio/turn.mp3',
    'follow': 'audio/follow.mp3', 'cang': 'audio/cang.mp3', 'cang25': 'audio/cang25.mp3',
    'win': 'audio/win.mp3', 'lost': 'audio/lost.mp3', '60': 'audio/60.mp3', 'no': 'audio/no.mp3',
    'knock': 'audio/knock.mp3', 'special': 'audio/special.mp3' // เพิ่มเสียง knock และ special
};

function resetAudioBuffersForContextRecovery() {
    Object.keys(audioBuffers).forEach(key => {
        delete audioBuffers[key];
        if (soundLoadStatus[key] === 'loaded') {
            soundLoadStatus[key] = 'idle';
        }
    });
    if (bgmSource) {
        try { bgmSource.onended = null; bgmSource.stop(); } catch (e) {}
        try { bgmSource.disconnect(); } catch (e) {}
        bgmSource = null;
    }
}

function getOrCreateAudioContext() {
    if (audioCtx && audioCtx.state === 'closed') {
        audioCtx = null;
        audioResumePromise = null;
        isAudioUnlocked = false;
        resetAudioBuffersForContextRecovery();
    }
    if (!audioCtx) {
        const AudioContextClass = window.AudioContext || window.webkitAudioContext;
        if (AudioContextClass) {
            try {
                audioCtx = new AudioContextClass();
                audioCtx.onstatechange = () => {
                    if (!audioCtx) return;
                    if (audioCtx.state === 'running') {
                        isAudioUnlocked = true;
                        processPendingSoundQueue();
                        ensureBgmPlayingIfWanted();
                    } else if (audioCtx.state === 'suspended' || audioCtx.state === 'interrupted') {
                        isAudioUnlocked = false;
                    } else if (audioCtx.state === 'closed') {
                        audioCtx = null;
                        audioResumePromise = null;
                        isAudioUnlocked = false;
                        resetAudioBuffersForContextRecovery();
                    }
                };
            } catch (e) {
                console.warn("Could not create AudioContext", e);
            }
        }
    }
    return audioCtx;
}

getOrCreateAudioContext();

function playSilentUnlockSample(ctx) {
    if (!ctx || ctx.state === 'closed') return;
    try {
        const sampleRate = (ctx.sampleRate && ctx.sampleRate > 0) ? ctx.sampleRate : 22050;
        const silentBuffer = ctx.createBuffer(1, 1, sampleRate);
        const source = ctx.createBufferSource();
        source.buffer = silentBuffer;
        source.connect(ctx.destination);
        if (typeof source.start === 'function') {
            source.start(0);
        } else if (typeof source.noteOn === 'function') {
            source.noteOn(0);
        }
    } catch (e) {}
}

function ensureAudioContextReady(fromUserGesture = false) {
    let ctx = getOrCreateAudioContext();
    if (!ctx) return Promise.resolve(false);
    if (ctx.state === 'running') {
        isAudioUnlocked = true;
        return Promise.resolve(true);
    }

    if (fromUserGesture || !audioResumePromise || ctx.state === 'suspended' || ctx.state === 'interrupted') {
        try {
            if (fromUserGesture) {
                playSilentUnlockSample(ctx);
            }
            const p = ctx.resume();
            if (p && typeof p.then === 'function') {
                const trackedPromise = p.then(() => {
                    if (audioResumePromise === trackedPromise) {
                        audioResumePromise = null;
                    }
                    const running = (ctx.state === 'running');
                    if (running) {
                        isAudioUnlocked = true;
                        processPendingSoundQueue();
                        ensureBgmPlayingIfWanted();
                    }
                    return running;
                }).catch(() => {
                    if (audioResumePromise === trackedPromise) {
                        audioResumePromise = null;
                    }
                    return ctx.state === 'running';
                });
                audioResumePromise = trackedPromise;
            } else {
                const running = (ctx.state === 'running');
                if (running) isAudioUnlocked = true;
                audioResumePromise = null;
                return Promise.resolve(running);
            }
        } catch (e) {
            audioResumePromise = null;
            return Promise.resolve(ctx.state === 'running');
        }
    }
    return audioResumePromise || Promise.resolve(ctx.state === 'running');
}

function decodeAudioDataAsync(ctx, arrayBuffer) {
    return new Promise((resolve, reject) => {
        if (!ctx || !arrayBuffer || arrayBuffer.byteLength === 0) {
            reject(new Error("Invalid AudioContext or empty ArrayBuffer"));
            return;
        }
        let settled = false;
        const safeResolve = (decoded) => {
            if (!settled) {
                settled = true;
                resolve(decoded);
            }
        };
        const safeReject = (err) => {
            if (!settled) {
                settled = true;
                reject(err || new Error("decodeAudioData error"));
            }
        };
        try {
            const clonedBuffer = arrayBuffer.slice(0);
            const res = ctx.decodeAudioData(
                clonedBuffer,
                (decoded) => safeResolve(decoded),
                (err) => safeReject(err)
            );
            if (res && typeof res.then === 'function') {
                res.then(safeResolve).catch(safeReject);
            }
        } catch (e) {
            safeReject(e);
        }
    });
}

async function fetchSoundArrayBuffer(url) {
    let lastErr = null;
    for (let attempt = 1; attempt <= MAX_SOUND_RETRY; attempt++) {
        try {
            const response = await fetch(url);
            if (!response.ok) throw new Error(`HTTP status ${response.status}`);
            const arrayBuffer = await response.arrayBuffer();
            if (!arrayBuffer || arrayBuffer.byteLength === 0) {
                throw new Error("Empty audio file buffer");
            }
            return arrayBuffer;
        } catch (e) {
            lastErr = e;
        }
    }
    throw lastErr || new Error("Failed to fetch sound file");
}

function loadSingleSound(key, forceReload = false) {
    if (!soundFiles[key]) {
        soundLoadStatus[key] = 'failed';
        return Promise.resolve(null);
    }
    const currentCtx = getOrCreateAudioContext();
    if (!forceReload && soundLoadStatus[key] === 'loaded' && audioBuffers[key] && currentCtx && currentCtx.state !== 'closed') {
        return Promise.resolve(audioBuffers[key]);
    }
    if (soundLoadStatus[key] === 'loading' && soundLoadPromises[key]) {
        return soundLoadPromises[key];
    }

    soundLoadStatus[key] = 'loading';
    soundLoadPromises[key] = (async () => {
        try {
            let rawBuffer = rawAudioBuffers[key];
            if (!rawBuffer || rawBuffer.byteLength === 0 || forceReload) {
                rawBuffer = await fetchSoundArrayBuffer(soundFiles[key]);
                rawAudioBuffers[key] = rawBuffer;
            }

            let ctx = getOrCreateAudioContext();
            if (!ctx) throw new Error("AudioContext unavailable");

            let audioBuffer = null;
            try {
                audioBuffer = await decodeAudioDataAsync(ctx, rawBuffer);
            } catch (decodeErr) {
                rawBuffer = await fetchSoundArrayBuffer(soundFiles[key]);
                rawAudioBuffers[key] = rawBuffer;
                ctx = getOrCreateAudioContext();
                if (!ctx) throw new Error("AudioContext unavailable");
                audioBuffer = await decodeAudioDataAsync(ctx, rawBuffer);
            }

            if (!audioBuffer) throw new Error("Decoded buffer is null");

            audioBuffers[key] = audioBuffer;
            soundLoadStatus[key] = 'loaded';
            soundRetryCounts[key] = 0;
            return audioBuffer;
        } catch (e) {
            soundRetryCounts[key] = (soundRetryCounts[key] || 0) + 1;
            soundLoadStatus[key] = 'failed';
            console.warn("Could not load sound", key, e);
            return null;
        } finally {
            delete soundLoadPromises[key];
            if (audioCtx && audioCtx.state === 'running') {
                processPendingSoundQueue();
                if (key === 'bgm') ensureBgmPlayingIfWanted();
            }
        }
    })();

    return soundLoadPromises[key];
}

async function loadSounds() {
    const keys = Object.keys(soundFiles);
    await Promise.allSettled(keys.map(key => loadSingleSound(key)));
}
loadSounds();

function startBufferSourceNow(key, loop = false) {
    const ctx = getOrCreateAudioContext();
    if (!ctx || !audioBuffers[key] || ctx.state !== 'running') return null;
    try {
        const source = ctx.createBufferSource();
        source.buffer = audioBuffers[key];
        source.connect(ctx.destination);
        source.loop = loop;
        source.onended = () => {
            try { source.disconnect(); } catch (e) {}
            if (loop && bgmSource === source && isBgmWanted) {
                bgmSource = null;
                ensureBgmPlayingIfWanted();
            }
        };
        source.start(0);
        return source;
    } catch (e) {
        console.warn("Error starting sound source", key, e);
        return null;
    }
}

function ensureBgmPlayingIfWanted() {
    if (!isBgmWanted) return;
    if (bgmSource) return;
    const ctx = getOrCreateAudioContext();
    if (!ctx || ctx.state !== 'running') return;
    if (soundLoadStatus['bgm'] === 'loaded' && audioBuffers['bgm']) {
        bgmSource = startBufferSourceNow('bgm', true);
    } else if (soundLoadStatus['bgm'] !== 'loading') {
        if (soundLoadStatus['bgm'] !== 'failed' || (soundRetryCounts['bgm'] || 0) < MAX_SOUND_RETRY) {
            loadSingleSound('bgm', soundLoadStatus['bgm'] === 'failed');
        }
    }
}

async function processPendingSoundQueue() {
    if (isProcessingSoundQueue) return;
    if (pendingSoundQueue.length === 0) return;

    const ctx = getOrCreateAudioContext();
    if (!ctx) return;

    if (ctx.state !== 'running') {
        ensureAudioContextReady(false);
        return;
    }

    isProcessingSoundQueue = true;
    try {
        while (pendingSoundQueue.length > 0) {
            const currentCtx = getOrCreateAudioContext();
            if (!currentCtx || currentCtx.state !== 'running') {
                ensureAudioContextReady(false);
                break;
            }

            const item = pendingSoundQueue[0];
            const { key, loop } = item;

            if (!soundFiles[key]) {
                pendingSoundQueue.shift();
                continue;
            }

            if (soundLoadStatus[key] !== 'loaded' || !audioBuffers[key]) {
                if (soundLoadStatus[key] === 'failed' && (soundRetryCounts[key] || 0) >= MAX_SOUND_RETRY) {
                    pendingSoundQueue.shift();
                    continue;
                }
                const buffer = await loadSingleSound(key, soundLoadStatus[key] === 'failed');
                if (!buffer) {
                    if ((soundRetryCounts[key] || 0) >= MAX_SOUND_RETRY) {
                        pendingSoundQueue.shift();
                        continue;
                    }
                    break;
                }
                const afterLoadCtx = getOrCreateAudioContext();
                if (!afterLoadCtx || afterLoadCtx.state !== 'running') {
                    ensureAudioContextReady(false);
                    break;
                }
            }

            pendingSoundQueue.shift();

            if (key === 'bgm' && loop) {
                if (isBgmWanted) {
                    if (bgmSource) {
                        try { bgmSource.onended = null; bgmSource.stop(); } catch (e) {}
                        try { bgmSource.disconnect(); } catch (e) {}
                        bgmSource = null;
                    }
                    const startedBgm = startBufferSourceNow('bgm', true);
                    if (startedBgm) {
                        bgmSource = startedBgm;
                    } else {
                        const checkCtx = getOrCreateAudioContext();
                        if (!checkCtx || checkCtx.state !== 'running') {
                            pendingSoundQueue.unshift(item);
                            ensureAudioContextReady(false);
                            break;
                        }
                    }
                }
            } else {
                const started = startBufferSourceNow(key, loop);
                if (!started) {
                    const checkCtx = getOrCreateAudioContext();
                    if (!checkCtx || checkCtx.state !== 'running') {
                        pendingSoundQueue.unshift(item);
                        ensureAudioContextReady(false);
                        break;
                    }
                }
            }
        }
    } finally {
        isProcessingSoundQueue = false;
        if (pendingSoundQueue.length > 0 && audioCtx && audioCtx.state === 'running') {
            const nextKey = pendingSoundQueue[0].key;
            if (soundLoadStatus[nextKey] === 'loaded' && audioBuffers[nextKey]) {
                processPendingSoundQueue();
            }
        }
    }
}

function playSound(key, loop = false) {
    if (!key || !soundFiles[key]) return null;

    if (!loop) {
        const now = Date.now();
        const lastTime = recentPlayedSignatureMap.get(key) || 0;
        if (now - lastTime < 35) {
            return null;
        }
        recentPlayedSignatureMap.set(key, now);
    }

    const ctx = getOrCreateAudioContext();
    if (!ctx) return null;

    if (ctx.state !== 'running') {
        ensureAudioContextReady(false);
    }

    const isBufferReady = (soundLoadStatus[key] === 'loaded' && !!audioBuffers[key]);
    const isContextRunning = (ctx.state === 'running');

    if (isContextRunning && isBufferReady && pendingSoundQueue.length === 0) {
        const source = startBufferSourceNow(key, loop);
        if (source) return source;
    }

    if (soundLoadStatus[key] !== 'loaded' && soundLoadStatus[key] !== 'loading') {
        if (soundLoadStatus[key] === 'failed') {
            soundRetryCounts[key] = 0;
        }
        loadSingleSound(key, soundLoadStatus[key] === 'failed');
    }

    if (key === 'bgm' && loop) {
        const hasBgmQueued = pendingSoundQueue.some(item => item.key === 'bgm' && item.loop);
        if (!hasBgmQueued) {
            pendingSoundQueue.push({ key, loop: true, timestamp: Date.now() });
        }
    } else {
        pendingSoundQueue.push({ key, loop, timestamp: Date.now() });
    }

    ensureAudioContextReady(false).then((ready) => {
        if (ready || (audioCtx && audioCtx.state === 'running')) {
            processPendingSoundQueue();
            ensureBgmPlayingIfWanted();
        }
    });

    return null;
}

function stopBGM() {
    isBgmWanted = false;
    for (let i = pendingSoundQueue.length - 1; i >= 0; i--) {
        if (pendingSoundQueue[i].key === 'bgm' && pendingSoundQueue[i].loop) {
            pendingSoundQueue.splice(i, 1);
        }
    }
    if (bgmSource) {
        try { bgmSource.onended = null; bgmSource.stop(); } catch(e) {}
        try { bgmSource.disconnect(); } catch(e) {}
        bgmSource = null;
    }
}

function playBGM() {
    stopBGM();
    isBgmWanted = true;
    bgmSource = playSound('bgm', true);
}

function unlockAudioOnUserGesture() {
    const ctx = getOrCreateAudioContext();
    if (!ctx) return;

    if (!isAudioUnlocked || ctx.state !== 'running') {
        playSilentUnlockSample(ctx);
    }

    if (ctx.state !== 'running') {
        ensureAudioContextReady(true).then((ready) => {
            if (ready || (audioCtx && audioCtx.state === 'running')) {
                isAudioUnlocked = true;
                processPendingSoundQueue();
                ensureBgmPlayingIfWanted();
            }
        });
    } else {
        isAudioUnlocked = true;
        if (pendingSoundQueue.length > 0 || (isBgmWanted && !bgmSource)) {
            processPendingSoundQueue();
            ensureBgmPlayingIfWanted();
        }
    }
}

['click', 'touchstart', 'touchend', 'keydown', 'keyup', 'pointerdown', 'pointerup', 'mousedown', 'mouseup'].forEach(evtName => {
    window.addEventListener(evtName, unlockAudioOnUserGesture, { capture: true, passive: true, once: false });
    document.addEventListener(evtName, unlockAudioOnUserGesture, { capture: true, passive: true, once: false });
});

function handleAudioStateRestore() {
    if (document.visibilityState === 'hidden') return;
    const ctx = getOrCreateAudioContext();
    if (!ctx) return;
    if (ctx.state !== 'running') {
        ensureAudioContextReady(false).then((ready) => {
            if (ready || (audioCtx && audioCtx.state === 'running')) {
                processPendingSoundQueue();
                ensureBgmPlayingIfWanted();
            }
        });
    } else {
        processPendingSoundQueue();
        ensureBgmPlayingIfWanted();
    }
}

document.addEventListener('visibilitychange', handleAudioStateRestore, { passive: true });
window.addEventListener('pageshow', handleAudioStateRestore, { passive: true });
window.addEventListener('focus', handleAudioStateRestore, { passive: true });
window.addEventListener('online', () => {
    Object.keys(soundFiles).forEach(key => {
        if (soundLoadStatus[key] === 'failed') {
            soundRetryCounts[key] = 0;
            loadSingleSound(key, true);
        }
    });
    handleAudioStateRestore();
}, { passive: true });

function sendSoundMessageToConn(conn, payload) {
    if (!conn) return;
    if (conn.open) {
        try {
            conn.send(payload);
        } catch (e) {
            console.warn("Failed to send sound message to peer", conn.peer, e);
        }
    } else {
        if (!pendingPeerSoundPayloads.has(conn.peer)) {
            pendingPeerSoundPayloads.set(conn.peer, []);
            conn.on('open', () => {
                const queued = pendingPeerSoundPayloads.get(conn.peer) || [];
                pendingPeerSoundPayloads.delete(conn.peer);
                queued.forEach(msg => {
                    try {
                        if (conn.open) conn.send(msg);
                    } catch (e) {}
                });
            });
        }
        pendingPeerSoundPayloads.get(conn.peer).push(payload);
    }
}

function handleIncomingPlaySoundData(data) {
    if (!data || !data.sound) return;
    if (isHost) return; // ป้องกันเสียงซ้ำในเครื่อง Host กรณีได้รับ Broadcast ย้อนกลับ
    if (data.soundEventId) {
        if (lastHandledRemoteSoundId === data.soundEventId) return;
        lastHandledRemoteSoundId = data.soundEventId;
    }
    playSound(data.sound);
}

function triggerSound(soundKey) {
    if (!soundKey) return;
    const soundEventId = `${myPeerId || 'host'}_${Date.now()}_${++soundEventSeq}`;
    if (myPeerId) playSound(soundKey);
    const payload = { type: 'PLAY_SOUND', sound: soundKey, soundEventId };
    guestConnections.forEach(c => {
        if (c && c.peer !== myPeerId) {
            sendSoundMessageToConn(c, payload);
        }
    });
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

const BOT_NAMES_POOL = [
    "บอทเมเปิ้ล",
    "บอทตาหวาน",
    "บอทข้าวปั้น",
    "บอทโมจิ",
    "บอทถุงเงิน",
    "บอทมะลิ",
    "บอทไข่มุก",
    "บอทสายไหม",
    "บอทโกโก้",
    "บอทเจ้าถั่ว",
    "บอทใบชา",
    "บอทปุยนุ่น",
    "บอทน้ำผึ้ง",
    "บอทคุกกี้",
    "บอทลูกพลับ",
    "บอทข้าวเหนียว",
    "บอทฟองนม",
    "บอทปลาทู",
    "บอทจุ๊บจิ๊บ",
    "บอทเจ้าขนม"
];

const BOT_NAMES_POOL_EN = [
    "Bot Maple",
    "Bot Sweetie",
    "Bot Onigiri",
    "Bot Mochi",
    "Bot Moneybag",
    "Bot Jasmine",
    "Bot Boba",
    "Bot Cotton",
    "Bot Cocoa",
    "Bot Peanut",
    "Bot Matcha",
    "Bot Fluffy",
    "Bot Honey",
    "Bot Cookie",
    "Bot Persimmon",
    "Bot StickyRice",
    "Bot Milkfoam",
    "Bot Mackerel",
    "Bot Chirpy",
    "Bot Muffin"
];

let peer = null, myPeerId = null, roomId = null, isHost = false;
let hostConnection = null, guestConnections = []; 
let roomPlayers = [], botColors = [], botNames = [], currentBotCount = 0, selectedColorId = 'red';
let myPlayerName = "", isNameConfirmed = false;
let isSystemInitialized = false;
let systemInitPromise = null;
let deck = [], discardPile = [];
let gameState = { turnIndex: 0, players: [], status: 'WAITING', topCardOwnerId: null, flowSourceId: null, skipPreVotes: [] };
let localPlayerState = { hand: [], hasDrawn: false, hasDiscarded: false, discardedRank: null, points: 200, turnCount: 0, hasFinishedFirstTurn: false, hasFlowedThisTurn: false, temporaryBot: false, returnRequested: false };
let globalPlayersMap = {}; 
let turnTimerInterval = null;
let announcerTimeout = null;
let hostHeartbeatTimer = null;
let heartbeatInterval = null;

// Timers for Accessibility Focus & Auto End Turn
let autoDrawFocusTimeout = null;
let autoHandFocusTimeout = null;
let autoEndTurnTimeout = null;

function clearAutoDrawFocusTimer() {
    if (autoDrawFocusTimeout) {
        clearTimeout(autoDrawFocusTimeout);
        autoDrawFocusTimeout = null;
    }
}

function clearAutoHandFocusTimer() {
    if (autoHandFocusTimeout) {
        clearTimeout(autoHandFocusTimeout);
        autoHandFocusTimeout = null;
    }
}

function clearAutoEndTurnTimer() {
    if (autoEndTurnTimeout) {
        clearTimeout(autoEndTurnTimeout);
        autoEndTurnTimeout = null;
    }
}

function scheduleFocusOnDrawButton() {
    clearAutoDrawFocusTimer();
    autoDrawFocusTimeout = setTimeout(() => {
        autoDrawFocusTimeout = null;
        if (isLocalTempBotModalOpen) return;
        const pub = window.currentPublicState;
        if (!pub || pub.status !== 'PLAYING' || pub.turnId !== myPeerId) return;
        if (localPlayerState.temporaryBot || localPlayerState.hasDrawn || localPlayerState.hasFlowedThisTurn) return;
        const btnDraw = document.getElementById('btn-draw');
        if (btnDraw && !btnDraw.disabled && btnDraw.style.display !== 'none') {
            btnDraw.focus();
        }
    }, 3000);
}

function scheduleFocusOnFirstHandCard() {
    clearAutoHandFocusTimer();
    autoHandFocusTimeout = setTimeout(() => {
        autoHandFocusTimeout = null;
        if (isLocalTempBotModalOpen) return;
        const pub = window.currentPublicState;
        if (!pub || pub.status !== 'PLAYING' || pub.turnId !== myPeerId) return;
        if (localPlayerState.temporaryBot || !localPlayerState.hasDrawn || localPlayerState.hasFlowedThisTurn) return;
        const handUi = document.getElementById('my-hand-ui');
        if (handUi && handUi.firstElementChild && !handUi.firstElementChild.disabled) {
            handUi.firstElementChild.focus();
        }
    }, 2100);
}

function scheduleAutoEndTurn(expectedPlayerId) {
    clearAutoEndTurnTimer();
    autoEndTurnTimeout = setTimeout(() => {
        autoEndTurnTimeout = null;
        if (!isHost || gameState.status !== 'PLAYING' || isEndingTurn) return;
        let currentCp = gameState.players[gameState.turnIndex];
        if (!currentCp || currentCp.id !== expectedPlayerId || currentCp.isBot) return;
        if (!(currentCp.hasDiscardedTurn || currentCp.hasFlowedThisTurn)) return;
        processPlayerAction(expectedPlayerId, 'END_TURN', 0, true);
    }, 1100);
}

// =========================================================
// Deferred Initialization Management (Firebase & PeerJS)
// =========================================================
function initFirebaseAndPeerSystem() {
    if (isSystemInitialized) return Promise.resolve();
    if (systemInitPromise) return systemInitPromise;

    systemInitPromise = new Promise((resolve, reject) => {
        try {
            if (!myPeerId) {
                myPeerId = "cang-client-" + Math.random().toString(36).substring(2, 8);
            }
            if (!peer) {
                peer = new Peer(myPeerId);
            }

            const roomsRef = ref(db, GAME_DB_ID + '/rooms');
            onValue(roomsRef, (snapshot) => {
                roomsSnapshotCache = snapshot.exists() ? snapshot.val() : {};
                renderRoomListFromCache();
            }, (err) => {
                console.warn("Firebase onValue error", err);
            });

            isSystemInitialized = true;
            resolve();
        } catch (e) {
            console.error("Failed to initialize Firebase and PeerJS", e);
            reject(e);
        }
    });

    return systemInitPromise;
}

// =========================================================
// Timeout, Temporary Bot & Disconnect States
// =========================================================
let currentTurnMaxTime = 40;
let isProcessingTurnTimeout = false;
let isEndingTurn = false;
let activeBotSequencePlayerId = null;
let botResponseTimers = {}; // Host map: { [playerId]: { interval, timeLeft, responded } }
let localBotModalTimer = null;
let hasSubmittedBotModalResponse = false;
let isLocalTempBotModalOpen = false;
let processedDisconnects = new Set();
let isClosingHostRoom = false;

function randomizeAllBotNames(count) {
    let shuffled = [...BOT_NAMES_POOL].sort(() => Math.random() - 0.5);
    botNames = shuffled.slice(0, count);
}

function triggerHostDisconnect() {
    clearTimeout(hostHeartbeatTimer);
    clearAutoDrawFocusTimer();
    clearAutoHandFocusTimer();
    clearAutoEndTurnTimer();
    if (localBotModalTimer) {
        clearInterval(localBotModalTimer);
        localBotModalTimer = null;
    }
    closeTemporaryBotModal();

    const nameScreen = document.getElementById('name-entry-screen');
    if (nameScreen) nameScreen.style.display = 'none';
    document.getElementById('setup-screen').style.display = 'none';
    document.getElementById('waiting-room-screen').style.display = 'none';
    document.getElementById('game-screen').style.display = 'none';
    document.getElementById('result-modal').style.display = 'none';
    document.getElementById('countdown-screen').style.display = 'none';
    document.getElementById('kang-animation-screen').style.display = 'none';
    document.getElementById('transition-overlay').style.display = 'none';
    
    const dcScreen = document.getElementById('disconnect-screen');
    dcScreen.style.display = 'flex';
    updateHomeBtnVisibility();
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

// =========================================================
// Temporary Bot Modal & Accessibility Management
// =========================================================
function setBackgroundInertForBotModal(isInert) {
    const bgIds = ['btn-home-top', 'btn-lang-toggle', 'game-screen', 'result-modal', 'waiting-room-screen', 'setup-screen', 'manual-screen', 'name-entry-screen'];
    bgIds.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        if (isInert) {
            el.setAttribute('aria-hidden', 'true');
            el.setAttribute('inert', '');
        } else {
            el.removeAttribute('aria-hidden');
            el.removeAttribute('inert');
        }
    });
}

function updateTemporaryBotModalCountdownText(seconds) {
    const cdEl = document.getElementById('temporary-bot-countdown');
    if (cdEl) {
        cdEl.innerText = t('tempBotCountdown', seconds);
    }
}

function openTemporaryBotModal(initialSeconds = 30) {
    const modal = document.getElementById('temporary-bot-modal');
    if (!modal) return;

    hasSubmittedBotModalResponse = false;
    isLocalTempBotModalOpen = true;

    const btnYes = document.getElementById('btn-bot-return-yes');
    const btnNo = document.getElementById('btn-bot-return-no');
    if (btnYes) btnYes.disabled = false;
    if (btnNo) btnNo.disabled = false;

    updateTemporaryBotModalCountdownText(initialSeconds);
    setBackgroundInertForBotModal(true);
    modal.style.display = 'flex';

    if (localBotModalTimer) clearInterval(localBotModalTimer);
    let localSec = initialSeconds;
    localBotModalTimer = setInterval(() => {
        localSec--;
        if (localSec >= 0) {
            updateTemporaryBotModalCountdownText(localSec);
        }
        if (localSec <= 0) {
            clearInterval(localBotModalTimer);
            localBotModalTimer = null;
        }
    }, 1000);

    setTimeout(() => {
        const heading = document.getElementById('temporary-bot-heading');
        if (heading) heading.focus();
    }, 100);
}

function closeTemporaryBotModal() {
    const modal = document.getElementById('temporary-bot-modal');
    if (localBotModalTimer) {
        clearInterval(localBotModalTimer);
        localBotModalTimer = null;
    }
    isLocalTempBotModalOpen = false;
    if (modal) {
        modal.style.display = 'none';
    }
    setBackgroundInertForBotModal(false);
}

function setupTemporaryBotModalEvents() {
    const btnYes = document.getElementById('btn-bot-return-yes');
    const btnNo = document.getElementById('btn-bot-return-no');

    if (btnYes) {
        btnYes.onclick = () => handleBotModalChoice(true);
    }
    if (btnNo) {
        btnNo.onclick = () => handleBotModalChoice(false);
    }
}

function handleBotModalChoice(wantsToReturn) {
    if (hasSubmittedBotModalResponse) return;
    hasSubmittedBotModalResponse = true;

    const btnYes = document.getElementById('btn-bot-return-yes');
    const btnNo = document.getElementById('btn-bot-return-no');
    if (btnYes) btnYes.disabled = true;
    if (btnNo) btnNo.disabled = true;

    closeTemporaryBotModal();

    if (wantsToReturn) {
        if (isHost) {
            handlePlayerBotReturnYes(myPeerId);
        } else if (hostConnection && hostConnection.open) {
            hostConnection.send({ type: 'BOT_RETURN_YES', playerId: myPeerId });
        }
    } else {
        if (isHost) {
            handleHostRoomExit();
        } else {
            if (hostConnection && hostConnection.open) {
                try {
                    hostConnection.send({ type: 'CLIENT_LEAVE_ROOM', playerId: myPeerId });
                } catch (e) {}
            }
            setTimeout(() => {
                try { if (hostConnection) hostConnection.close(); } catch (e) {}
                try { if (peer) peer.destroy(); } catch (e) {}
                window.location.href = '../index.html';
            }, 150);
        }
    }
}

function clearBotResponseTimer(playerId) {
    if (botResponseTimers[playerId]) {
        clearInterval(botResponseTimers[playerId].interval);
        delete botResponseTimers[playerId];
    }
}

function startHostBotResponseTimer(playerId) {
    clearBotResponseTimer(playerId);
    let stateObj = {
        timeLeft: 30,
        responded: false,
        interval: null
    };
    botResponseTimers[playerId] = stateObj;

    if (playerId === myPeerId) {
        openTemporaryBotModal(30);
    } else {
        let conn = guestConnections.find(c => c.peer === playerId);
        if (conn) {
            conn.send({ type: 'SHOW_TEMP_BOT_MODAL', timeLeft: 30 });
        }
    }

    stateObj.interval = setInterval(() => {
        if (stateObj.responded) {
            clearBotResponseTimer(playerId);
            return;
        }
        stateObj.timeLeft--;

        if (playerId === myPeerId) {
            if (stateObj.timeLeft >= 0) updateTemporaryBotModalCountdownText(stateObj.timeLeft);
        } else {
            let conn = guestConnections.find(c => c.peer === playerId);
            if (conn && stateObj.timeLeft >= 0) {
                conn.send({ type: 'TEMP_BOT_MODAL_COUNTDOWN', timeLeft: stateObj.timeLeft });
            }
        }

        if (stateObj.timeLeft <= 0) {
            stateObj.responded = true;
            clearBotResponseTimer(playerId);
            handleBotModalTimeoutOnHost(playerId);
        }
    }, 1000);
}

function handleBotModalTimeoutOnHost(playerId) {
    if (playerId === myPeerId) {
        closeTemporaryBotModal();
        handleHostRoomExit();
    } else {
        let conn = guestConnections.find(c => c.peer === playerId);
        if (conn) {
            try {
                conn.send({ type: 'FORCE_LEAVE_TO_INDEX' });
            } catch (e) {}
        }
        handleClientDeparture(playerId, 'timeout');
    }
}

function handlePlayerBotReturnYes(playerId) {
    if (!isHost) return;
    let timerObj = botResponseTimers[playerId];
    if (timerObj) {
        if (timerObj.responded) return;
        timerObj.responded = true;
        clearBotResponseTimer(playerId);
    }

    let p = gameState.players.find(x => x.id === playerId);
    if (!p || p.permanentBot || !p.temporaryBot) return;
    if (p.returnRequested) return;

    p.returnRequested = true;

    if (playerId === myPeerId) {
        closeTemporaryBotModal();
    } else {
        let conn = guestConnections.find(c => c.peer === playerId);
        if (conn) {
            conn.send({ type: 'CLOSE_TEMP_BOT_MODAL' });
        }
    }

    syncStateToAll();
}

// =========================================================
// Unified Client Departure & Host Room Exit Handlers
// =========================================================
function handleClientDeparture(playerId, reason = 'disconnect') {
    if (!isHost || !playerId || playerId === myPeerId) return;
    if (processedDisconnects.has(playerId)) return;
    processedDisconnects.add(playerId);

    clearBotResponseTimer(playerId);
    pendingPeerSoundPayloads.delete(playerId);

    // 1. ปิดและลบ Connection ออกจาก guestConnections อย่างปลอดภัย
    let connIdx = guestConnections.findIndex(c => c.peer === playerId);
    if (connIdx !== -1) {
        let conn = guestConnections[connIdx];
        guestConnections.splice(connIdx, 1);
        try {
            conn.close();
        } catch (e) {}
    }

    // 2. กรณีอยู่ในห้องรอ (WAITING) ก่อนเริ่มเกม
    if (gameState.status === 'WAITING') {
        let rpIdx = roomPlayers.findIndex(p => p.id === playerId);
        let removedName = resolveName(playerId);
        if (rpIdx !== -1) {
            roomPlayers.splice(rpIdx, 1);
        }
        let joIdx = window.joinOrder.indexOf(playerId);
        if (joIdx !== -1) {
            window.joinOrder.splice(joIdx, 1);
        }
        broadcastLobby();
        broadcastAnnounce(`${removedName} ออกจากห้องรอผู้เล่น`);
        return;
    }

    // 3. กรณีอยู่ในเกม เปลี่ยนผู้เล่นคนนั้นเป็นบอทถาวร โดยรักษา Player ID, ชื่อ, ไพ่, คะแนน และลำดับเดิม
    let p = gameState.players.find(x => x.id === playerId);
    if (!p) return;

    p.isBot = true;
    p.originalIsBot = false;
    p.temporaryBot = false;
    p.returnRequested = false;
    p.permanentBot = true;

    broadcastAnnounce(`[PID:${p.id}] ออกจากห้อง ระบบเปลี่ยนเป็นบอทเล่นแทนจนจบเกม`);

    // หากอยู่ในช่วง PRE_GAME ให้ตรวจสอบว่าผู้เล่นจริงที่เหลือโหวตข้ามครบหรือยัง
    if (gameState.status === 'PRE_GAME') {
        syncStateToAll();
        let realPlayers = gameState.players.filter(x => !x.isOut && !x.isBot);
        if (realPlayers.length === 0 || (gameState.skipPreVotes && realPlayers.every(rp => gameState.skipPreVotes.includes(rp.id)))) {
            clearInterval(turnTimerInterval);
            triggerSound('60');
            gameState.status = 'PLAYING';
            let firstPlayer = gameState.players[gameState.turnIndex] || gameState.players[0];
            broadcastAnnounce(`📢 ผู้เล่นทุกคนพร้อมแล้ว! ขอเชิญ [PID:${firstPlayer.id}] จั่วไพ่ใบแรกและทิ้งลงกองเพื่อเริ่มเกมได้เลยครับ`);
            syncStateToAll();
            startTurnTimer();
            checkBotTurn();
        }
        return;
    }

    syncStateToAll();

    // หากเป็นตาของผู้เล่นที่เพิ่งออกในสถานะ PLAYING ให้บอทเข้าควบคุมต่อทันทีโดยเกมไม่ค้าง
    if (gameState.status === 'PLAYING') {
        let currentTurnPlayer = gameState.players[gameState.turnIndex];
        if (currentTurnPlayer && currentTurnPlayer.id === p.id) {
            clearInterval(turnTimerInterval);
            checkBotTurn();
        }
    }
}

function handleHostRoomExit() {
    if (!isHost || isClosingHostRoom) return;
    isClosingHostRoom = true;

    clearInterval(turnTimerInterval);
    clearInterval(heartbeatInterval);
    clearAutoDrawFocusTimer();
    clearAutoHandFocusTimer();
    clearAutoEndTurnTimer();
    if (localBotModalTimer) {
        clearInterval(localBotModalTimer);
        localBotModalTimer = null;
    }
    Object.keys(botResponseTimers).forEach(pid => clearBotResponseTimer(pid));

    if (roomId) {
        try {
            remove(ref(db, GAME_DB_ID + '/rooms/' + roomId));
        } catch (e) {}
    }

    guestConnections.forEach(c => {
        try {
            c.send({ type: 'HOST_CLOSED_ROOM' });
        } catch (e) {}
    });

    setTimeout(() => {
        guestConnections.forEach(c => {
            try { c.close(); } catch (e) {}
        });
        guestConnections = [];
        try {
            if (peer) peer.destroy();
        } catch (e) {}
        window.location.href = '../index.html';
    }, 200);
}

let currentTopCardText = t('topCardEmpty');

// --- CHEAT MODE BUFFERS ---
let cheatBuffer = "";
let cheatTimeout = null;

// Shortcut Keys & Cheat Code Event Listener
document.addEventListener('keydown', (e) => {
    if (isLocalTempBotModalOpen) return;

    // 1. Cheat Mode Logic
    if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'TEXTAREA') {
        if (e.key.length === 1 && !e.altKey && !e.ctrlKey && !e.metaKey) {
            cheatBuffer += e.key;
            clearTimeout(cheatTimeout);
            cheatTimeout = setTimeout(() => { cheatBuffer = ""; }, 2000);

            const cmd = cheatBuffer.toLowerCase();
            if (cmd.includes("ตอง")) { setCheatHand("ตอง"); cheatBuffer = ""; }
            else if (cmd.includes("ดอก") || cmd.includes("สี")) { setCheatHand("ดอก"); cheatBuffer = ""; }
            else if (cmd.includes("เรียง")) { setCheatHand("เรียง"); cheatBuffer = ""; }
            else if (cmd.includes("50")) { setCheatHand("50"); cheatBuffer = ""; }
            else if (cmd.includes("3a")) { setCheatHand("3a"); cheatBuffer = ""; }
        }
    }

    // 2. Shortcut Keys Logic
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
    const langBtn = document.getElementById('btn-lang-toggle');
    const nameScreen = document.getElementById('name-entry-screen');
    const setupScreen = document.getElementById('setup-screen');
    const nameDisp = nameScreen ? window.getComputedStyle(nameScreen).display : 'none';
    const setupDisp = setupScreen ? window.getComputedStyle(setupScreen).display : 'none';
    const isTopBarVisible = (setupDisp !== 'none' || nameDisp !== 'none');

    if (homeBtn) homeBtn.style.display = isTopBarVisible ? 'block' : 'none';
    if (langBtn) langBtn.style.display = isTopBarVisible ? 'block' : 'none';
}

const bubblesBg = document.getElementById('bg-bubbles');
for(let i=0; i<15; i++) {
    let b = document.createElement('div'); b.className = 'bubble';
    b.style.width = Math.random() * 20 + 10 + 'px'; b.style.height = b.style.width;
    b.style.left = Math.random() * 100 + '%';
    b.style.animationDuration = Math.random() * 4 + 4 + 's';
    bubblesBg.appendChild(b);
}

function localizeBotName(rawName, fallbackIdx = 0) {
    if (!rawName) return currentLang === 'en' ? (BOT_NAMES_POOL_EN[fallbackIdx % BOT_NAMES_POOL_EN.length] || 'Bot') : (BOT_NAMES_POOL[fallbackIdx % BOT_NAMES_POOL.length] || 'บอท');
    const poolIdx = BOT_NAMES_POOL.indexOf(rawName);
    if (poolIdx !== -1) {
        return currentLang === 'en' ? BOT_NAMES_POOL_EN[poolIdx] : BOT_NAMES_POOL[poolIdx];
    }
    return rawName;
}

function resolveName(id) {
    if (typeof id === 'string' && id.startsWith('bot_')) {
        let bIdx = parseInt(id.split('_')[1]);
        let rawBName = (!isNaN(bIdx) && botNames[bIdx]) ? botNames[bIdx] : (BOT_NAMES_POOL[(isNaN(bIdx) ? 0 : bIdx) % BOT_NAMES_POOL.length]);
        return localizeBotName(rawBName, isNaN(bIdx) ? 0 : bIdx);
    }
    let p = globalPlayersMap[id];
    if (p && p.name && String(p.name).trim() !== '') {
        const trimmed = String(p.name).trim();
        return p.isBot ? localizeBotName(trimmed, 0) : trimmed;
    }
    if (gameState && gameState.players) {
        let gp = gameState.players.find(x => x.id === id);
        if (gp && gp.name && String(gp.name).trim() !== '') {
            const trimmed = String(gp.name).trim();
            return gp.isBot ? localizeBotName(trimmed, 0) : trimmed;
        }
    }
    if (roomPlayers) {
        let rp = roomPlayers.find(x => x.id === id);
        if (rp && rp.name && String(rp.name).trim() !== '') {
            return String(rp.name).trim();
        }
    }
    if (id === myPeerId && myPlayerName) return myPlayerName;
    return currentLang === 'en' ? 'Player' : 'ผู้เล่น';
}

// =========================================================
// UI Transitions & Announcements
// =========================================================
window.showTransition = (msg, duration) => {
    const overlay = document.getElementById('transition-overlay');
    const text = document.getElementById('transition-text');
    text.innerText = localizeDynamicMessage(msg);
    overlay.style.display = 'flex';
    if (duration) setTimeout(() => { overlay.style.display = 'none'; }, duration);
};

window.hideTransition = () => { document.getElementById('transition-overlay').style.display = 'none'; };

function broadcastTransition(msg, duration) {
    if(myPeerId) window.showTransition(msg, duration);
    guestConnections.forEach(c => c.send({ type: 'SHOW_TRANSITION', msg, duration }));
}

const PRE_GAME_BASE_ANNOUNCEMENT = "เปิดรอบใหม่! มีใครจะแคง 25 ไหม?";

function buildLocalPreGameAnnouncement(baseMsg) {
    const hand = (localPlayerState && Array.isArray(localPlayerState.hand)) ? localPlayerState.hand : [];
    if (hand.length === 0) return localizeDynamicMessage(baseMsg);

    const handSum = hand.reduce((sum, c) => sum + (c.val || getCardValue(c.rank)), 0);
    const specialType = getSpecialWinType(hand, localPlayerState);
    const localizedBase = localizeDynamicMessage(baseMsg);
    return t('preGameSummary', localizedBase, handSum, specialType);
}

function announce(rawMsg, showOnBanner = true) {
    if (rawMsg === PRE_GAME_BASE_ANNOUNCEMENT) {
        const pub = window.currentPublicState;
        if (pub && pub.status === 'PRE_GAME' && localPlayerState && Array.isArray(localPlayerState.hand) && localPlayerState.hand.length > 0) {
            window._pendingPreGameBaseMsg = null;
            window._announcedPreGameRound = true;
            rawMsg = buildLocalPreGameAnnouncement(rawMsg);
        } else {
            window._pendingPreGameBaseMsg = rawMsg;
            window._announcedPreGameRound = false;
            return;
        }
    }

    if (typeof rawMsg === 'string' && (rawMsg.includes('จั่วได้ไพ่') || rawMsg.includes('drew '))) {
        scheduleFocusOnFirstHandCard();
    }

    let formattedMsg = localizeDynamicMessage(rawMsg);
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
    announce(t('manualOpenedAnnounce'), false);
};

window.closeManual = () => {
    document.getElementById('manual-screen').style.display = 'none';
    document.getElementById('setup-screen').style.display = 'flex';
    updateHomeBtnVisibility();
    document.getElementById('btn-open-manual').focus();
    announce(t('manualClosedAnnounce'), false);
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
// Player Name Entry Logic
// =========================================================
function setupNameEntryEvents() {
    const nameInput = document.getElementById('player-name-input');
    const confirmBtn = document.getElementById('btn-confirm-name');
    const nameScreen = document.getElementById('name-entry-screen');
    const setupScreen = document.getElementById('setup-screen');

    if (!nameInput || !confirmBtn || !nameScreen) return;

    const handleConfirmName = async (e) => {
        if (e) e.preventDefault();
        if (isNameConfirmed) return;

        const trimmedName = nameInput.value.trim();
        if (!trimmedName) {
            announce(t('errEmptyName'), false);
            nameInput.focus();
            return;
        }

        isNameConfirmed = true;
        myPlayerName = trimmedName;
        confirmBtn.disabled = true;
        nameInput.disabled = true;

        playSound('select');

        try {
            await initFirebaseAndPeerSystem();
        } catch (err) {
            console.error("Error during initial services setup:", err);
        }

        nameScreen.style.display = 'none';
        if (setupScreen) {
            setupScreen.style.display = 'flex';
        }
        updateHomeBtnVisibility();
        announce(t('welcomeLobbyAnnounce', myPlayerName), false);

        setTimeout(() => {
            const createRoomBtn = document.getElementById('btn-create-room');
            if (createRoomBtn) createRoomBtn.focus();
        }, 100);
    };

    confirmBtn.addEventListener('click', handleConfirmName);
    nameInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            handleConfirmName(e);
        }
    });
}

// =========================================================
// Network & Lobby Logic
// =========================================================
function initializePageUI() {
    const langBtn = document.getElementById('btn-lang-toggle');
    if (langBtn) {
        langBtn.addEventListener('click', window.toggleLanguage);
    }
    applyLanguage(false);
    setupNameEntryEvents();
    setupTemporaryBotModalEvents();
    updateHomeBtnVisibility();
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializePageUI);
} else {
    initializePageUI();
}

window.createNewRoom = async () => {
    if (!myPlayerName || !myPlayerName.trim()) return;
    await initFirebaseAndPeerSystem();
    isHost = true; announce(t('creatingRoomAnnounce'), false);
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
                roomPlayers = [{ id: myPeerId, name: myPlayerName, color: selectedColorId }];
                botColors = [];
                botNames = [];
                currentBotCount = 0;
                processedDisconnects.clear();
                const myRoomRef = ref(db, GAME_DB_ID + '/rooms/' + roomId);
                set(myRoomRef, { roomId: roomId, playerCount: 1, state: 'WAITING' });
                onDisconnect(myRoomRef).remove();
                document.getElementById('setup-screen').style.display = 'none';
                document.getElementById('waiting-room-screen').style.display = 'flex';
                updateHomeBtnVisibility();
                announce(t('roomCreatedAnnounce', roomId), false);
                playSound('1');
                updateLobbyUI(); listenHostConnections();

                heartbeatInterval = setInterval(() => { guestConnections.forEach(c => c.send({ type: 'HEARTBEAT' })); }, 15000);
            });
        }
    });
};

window.joinRoomDirect = async (targetRoomId) => {
    if (!myPlayerName || !myPlayerName.trim()) return;
    await initFirebaseAndPeerSystem();
    isHost = false; roomId = targetRoomId;
    announce(t('joiningRoomAnnounce', roomId), false);
    document.getElementById('setup-screen').style.display = 'none';
    document.getElementById('waiting-room-screen').style.display = 'flex';
    document.getElementById('host-bot-section').style.display = 'none';
    updateHomeBtnVisibility();
    
    if (!peer) return;
    hostConnection = peer.connect("cang-host-" + roomId);
    hostConnection.on('open', () => { 
        hostConnection.send({ type: 'JOIN_ROOM', name: myPlayerName }); 
        announce(t('joinedRoomAnnounce'), false); 
        resetHeartbeat();
    });
    hostConnection.on('data', (data) => {
        resetHeartbeat();
        if (data.type === 'HEARTBEAT') return;
        if (data.type === 'HOST_CLOSED_ROOM') {
            triggerHostDisconnect();
            return;
        }
        if (data.type === 'FORCE_LEAVE_TO_INDEX') {
            closeTemporaryBotModal();
            try { if (hostConnection) hostConnection.close(); } catch (e) {}
            try { if (peer) peer.destroy(); } catch (e) {}
            window.location.href = '../index.html';
            return;
        }
        if (data.type === 'SHOW_TEMP_BOT_MODAL') {
            openTemporaryBotModal(data.timeLeft || 30);
            return;
        }
        if (data.type === 'TEMP_BOT_MODAL_COUNTDOWN') {
            if (localBotModalTimer) {
                clearInterval(localBotModalTimer);
                localBotModalTimer = null;
            }
            updateTemporaryBotModalCountdownText(data.timeLeft);
            return;
        }
        if (data.type === 'CLOSE_TEMP_BOT_MODAL') {
            closeTemporaryBotModal();
            return;
        }
        if (data.type === 'LOBBY_UPDATE') { 
            roomPlayers = data.roomPlayers || [];
            currentBotCount = data.botCount || 0;
            botColors = data.botColors || [];
            botNames = data.botNames || [];
            if (data.joinOrder) window.joinOrder = data.joinOrder; 
            updateLobbyUI(); 
        }
        if (data.type === 'START_COUNTDOWN') { 
            if (!window.isStartingRound) { window.isStartingRound = true; doCountdownAndStart(); }
        }
        if (data.type === 'SYNC_STATE') renderClientGame(data.publicState, data.privateState);
        if (data.type === 'ANIM_SC_FLOAT') showFloatingScore(data.targetId, data.amt);
        if (data.type === 'SHOW_RESULT') showResultModal(data.title, data.detail, data.isEnd, data.winners, data.losers);
        if (data.type === 'ANNOUNCE') announce(data.msg);
        if (data.type === 'TIMER_UPDATE') window.updateTimerUI(data.timeLeft, data.maxTime);
        if (data.type === 'PLAY_SOUND') handleIncomingPlaySoundData(data);
        if (data.type === 'SHOW_TRANSITION') window.showTransition(data.msg, data.duration);
        if (data.type === 'HIDE_TRANSITION') window.hideTransition();
        if (data.type === 'KANG_ANIMATION') window.playKangAnimation(data.callerId, data.winType);
    });
    hostConnection.on('close', () => {
        triggerHostDisconnect();
    });
    hostConnection.on('error', () => {
        triggerHostDisconnect();
    });
};

function listenHostConnections() {
    if (!peer) return;
    peer.on('connection', (conn) => {
        conn.on('data', (data) => {
            if (data.type === 'JOIN_ROOM') {
                if (roomPlayers.length + currentBotCount >= 6) return;
                let usedColors = roomPlayers.map(p => p.color).concat(botColors);
                let avail = COLORS.filter(c => !usedColors.includes(c.id));
                let randomColor = avail.length > 0 ? avail[Math.floor(Math.random() * avail.length)].id : 'red';
                let joinedName = (data.name && String(data.name).trim() !== '') ? String(data.name).trim() : 'ผู้เล่น';
                
                roomPlayers.push({ id: conn.peer, name: joinedName, color: randomColor });
                window.joinOrder.push(conn.peer); 
                guestConnections.push(conn); 
                broadcastLobby();
                broadcastAnnounce(`${joinedName} เข้าร่วมห้องใหม่ รวม ${roomPlayers.length + currentBotCount} คน`);
                triggerSound('select');
            }
            if (data.type === 'PLAYER_ACTION') processPlayerAction(data.playerId, data.action, data.cardIndex);
            if (data.type === 'BOT_RETURN_YES') handlePlayerBotReturnYes(data.playerId || conn.peer);
            if (data.type === 'CLIENT_LEAVE_ROOM') handleClientDeparture(data.playerId || conn.peer, 'voluntary');
        });
        conn.on('close', () => {
            handleClientDeparture(conn.peer, 'disconnect');
        });
        conn.on('error', () => {
            handleClientDeparture(conn.peer, 'error');
        });
    });
}

function broadcastLobby() {
    let data = { type: 'LOBBY_UPDATE', roomPlayers, botCount: currentBotCount, botColors, botNames, joinOrder: window.joinOrder };
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
        randomizeAllBotNames(currentBotCount);
    } else if (delta < 0 && currentBotCount > 0) {
        currentBotCount--; botColors.pop(); 
        for (let i = window.joinOrder.length - 1; i >= 0; i--) {
            if (window.joinOrder[i].startsWith('bot_')) { window.joinOrder.splice(i, 1); break; }
        }
        randomizeAllBotNames(currentBotCount);
    }
    broadcastLobby(); triggerSound('select');
};

function updateLobbyUI() {
    document.getElementById('display-room-id').innerText = roomId;
    document.getElementById('bot-count-text').innerText = t('botCountUnit', currentBotCount);
    
    globalPlayersMap = {};
    roomPlayers.forEach(p => {
        let c = COLORS.find(x => x.id === p.color) || COLORS[0];
        let pName = (p.name && String(p.name).trim() !== '') ? String(p.name).trim() : (p.id === myPeerId && myPlayerName ? myPlayerName : 'ผู้เล่น');
        globalPlayersMap[p.id] = { isBot: false, name: pName, colorName: c ? c.name : 'แดง', colorHex: c ? c.hex : 'var(--red)' };
    });
    botColors.forEach((b, i) => {
        let c = COLORS.find(x => x.id === b) || COLORS[i % COLORS.length];
        let bName = botNames[i] || BOT_NAMES_POOL[i % BOT_NAMES_POOL.length];
        globalPlayersMap['bot_' + i] = { isBot: true, name: bName, colorName: c ? c.name : 'แดง', colorHex: c ? c.hex : 'var(--red)' };
    });

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
    guestConnections.forEach(c => sendSoundMessageToConn(c, { type: 'START_COUNTDOWN' }));
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
                            let botIdx = parseInt(id.split('_')[1]);
                            let bColor = botColors[botIdx] || COLORS[botIdx % COLORS.length].id;
                            let bName = botNames[botIdx] || BOT_NAMES_POOL[botIdx % BOT_NAMES_POOL.length];
                            let cInfo = COLORS.find(c => c.id === bColor) || COLORS[0];
                            gameState.players.push({
                                id: id,
                                name: bName,
                                colorInfo: cInfo,
                                isBot: true,
                                originalIsBot: true,
                                temporaryBot: false,
                                returnRequested: false,
                                permanentBot: false,
                                consecutiveTimeouts: 0,
                                nextTurnTimeLimit: 40,
                                points: 200,
                                hand: [],
                                isOut: false,
                                turnCount: 0,
                                hasFinishedFirstTurn: false,
                                hasFlowedThisTurn: false
                            });
                        } else {
                            let p = roomPlayers.find(x => x.id === id);
                            if (p) {
                                let cInfo = COLORS.find(c => c.id === p.color) || COLORS[0];
                                let pName = (p.name && String(p.name).trim() !== '') ? String(p.name).trim() : (p.id === myPeerId && myPlayerName ? myPlayerName : 'ผู้เล่น');
                                gameState.players.push({
                                    id: p.id,
                                    name: pName,
                                    colorInfo: cInfo,
                                    isBot: false,
                                    originalIsBot: false,
                                    temporaryBot: false,
                                    returnRequested: false,
                                    permanentBot: false,
                                    consecutiveTimeouts: 0,
                                    nextTurnTimeLimit: 40,
                                    points: 200,
                                    hand: [],
                                    isOut: false,
                                    turnCount: 0,
                                    hasFinishedFirstTurn: false,
                                    hasFlowedThisTurn: false
                                });
                            }
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
function restorePlayerIfReturnRequested(p) {
    if (!p) return false;
    if (p.temporaryBot && p.returnRequested && !p.permanentBot && !p.originalIsBot) {
        p.isBot = false;
        p.temporaryBot = false;
        p.returnRequested = false;
        p.consecutiveTimeouts = 0;
        p.nextTurnTimeLimit = 40;
        return true;
    }
    return false;
}

function startNewRound() {
    clearAutoEndTurnTimer();
    clearAutoDrawFocusTimer();
    clearAutoHandFocusTimer();

    if (window.nextRoundHostId) {
        let idx = gameState.players.findIndex(p => p.id === window.nextRoundHostId);
        if (idx > 0) {
            let newHost = gameState.players.splice(idx, 1)[0];
            gameState.players.unshift(newHost);
        }
    }

    deck = buildDeck(); discardPile = [];
    gameState.status = 'PRE_GAME'; gameState.turnIndex = 0; gameState.topCardOwnerId = null; gameState.flowSourceId = null; gameState.skipPreVotes = []; 
    isProcessingTurnTimeout = false;
    isEndingTurn = false;
    activeBotSequencePlayerId = null;
    
    gameState.players.forEach((p) => {
        if(p.points > 0) { 
            p.hand = deck.splice(0, 5);
            sortPlayerHand(p.hand, null, false, false, false);
            p.isOut = false; p.hasDrawnTurn = false; p.hasDiscardedTurn = false; p.hasFlowedThisTurn = false; p.discardedRankThisTurn = null; p.turnCount = 0; p.hasFinishedFirstTurn = false;
        } else { p.isOut = true; } 
    });
    
    syncStateToAll(); startPreGameTimer(); broadcastAnnounce(PRE_GAME_BASE_ANNOUNCEMENT);
}

function startPreGameTimer() {
    clearInterval(turnTimerInterval);
    currentTurnMaxTime = 40;
    let timeLeft = 40; broadcastTimer(timeLeft, 40);
    turnTimerInterval = setInterval(() => {
        timeLeft--; broadcastTimer(timeLeft, 40);
        if(timeLeft <= 0) { 
            clearInterval(turnTimerInterval); 
            if (gameState.status === 'PRE_GAME') {
                triggerSound('60'); gameState.status = 'PLAYING';
                let firstPlayer = gameState.players[gameState.turnIndex] || gameState.players[0];
                restorePlayerIfReturnRequested(firstPlayer);
                broadcastAnnounce(`📢 หมดเวลาดูไพ่ 40 วินาที ไม่มีใครกดแคง! ขอเชิญ [PID:${firstPlayer.id}] จั่วไพ่ใบแรกและทิ้งลงกองเพื่อเริ่มเกมได้เลยครับ`);
                syncStateToAll(); startTurnTimer(); checkBotTurn();
            }
        }
    }, 1000);
}

function startTurnTimer() {
    clearInterval(turnTimerInterval);
    let cp = gameState.players[gameState.turnIndex];
    let maxTime = (cp && !cp.isBot && cp.nextTurnTimeLimit === 20) ? 20 : 40;
    currentTurnMaxTime = maxTime;
    let timeLeft = maxTime;
    broadcastTimer(timeLeft, maxTime);
    turnTimerInterval = setInterval(() => {
        timeLeft--; broadcastTimer(timeLeft, maxTime);
        if(timeLeft <= 0) { clearInterval(turnTimerInterval); handleTurnTimeout(); }
    }, 1000);
}

function broadcastTimer(t, maxT = currentTurnMaxTime) {
    if(myPeerId) window.updateTimerUI(t, maxT);
    guestConnections.forEach(c => c.send({ type: 'TIMER_UPDATE', timeLeft: t, maxTime: maxT }));
}

window.updateTimerUI = (timeLeft, maxTime = 40) => {
    const timerBar = document.getElementById('timer-bar-fill');
    if(timerBar) {
        let validMax = (maxTime && maxTime > 0) ? maxTime : 40;
        let pct = Math.max(0, (timeLeft / validMax) * 100);
        timerBar.style.width = pct + '%'; timerBar.style.background = timeLeft <= 10 ? 'var(--red)' : 'var(--orange)';
    }
};

function handleTurnTimeout() {
    if (isProcessingTurnTimeout) return;
    let cp = gameState.players[gameState.turnIndex];
    if (!cp || cp.isBot || gameState.status !== 'PLAYING') return;

    clearAutoEndTurnTimer();
    isProcessingTurnTimeout = true;
    try {
        cp.consecutiveTimeouts = (cp.consecutiveTimeouts || 0) + 1;

        if (cp.consecutiveTimeouts === 1) {
            // ระดับที่ 1: หมดเวลาครั้งที่ 1
            cp.nextTurnTimeLimit = 40;
            broadcastAnnounce(`⏰ [PID:${cp.id}] หมดเวลาเล่น (ครั้งที่ 1/3) ระบบเล่นแทนอัตโนมัติ`);
            if (discardPile.length === 0 && !cp.hasDiscardedTurn) {
                if (!cp.hasDrawnTurn) processPlayerAction(cp.id, 'DRAW', 0, true);
                if (cp.hand.length > 0) processPlayerAction(cp.id, 'DISCARD', 0, true);
            } else {
                if (!cp.hasDrawnTurn) processPlayerAction(cp.id, 'DRAW', 0, true);
                if (!cp.hasDiscardedTurn && cp.hand.length > 0) processPlayerAction(cp.id, 'DISCARD', 0, true);
            }
            if (gameState.status === 'PLAYING') processPlayerAction(cp.id, 'END_TURN', 0, true);
        } else if (cp.consecutiveTimeouts === 2) {
            // ระดับที่ 2: หมดเวลาครั้งที่ 2 ติดต่อกัน ลดเวลาตาถัดไปเหลือ 20 วินาที
            cp.nextTurnTimeLimit = 20;
            broadcastAnnounce(`⏰ [PID:${cp.id}] หมดเวลาเล่น 2 ครั้งติดต่อกัน! ตาถัดไปจะลดเวลาเหลือ 20 วินาที`);
            if (discardPile.length === 0 && !cp.hasDiscardedTurn) {
                if (!cp.hasDrawnTurn) processPlayerAction(cp.id, 'DRAW', 0, true);
                if (cp.hand.length > 0) processPlayerAction(cp.id, 'DISCARD', 0, true);
            } else {
                if (!cp.hasDrawnTurn) processPlayerAction(cp.id, 'DRAW', 0, true);
                if (!cp.hasDiscardedTurn && cp.hand.length > 0) processPlayerAction(cp.id, 'DISCARD', 0, true);
            }
            if (gameState.status === 'PLAYING') processPlayerAction(cp.id, 'END_TURN', 0, true);
        } else if (cp.consecutiveTimeouts >= 3) {
            // ระดับที่ 3: หมดเวลาครบ 3 ครั้ง เปลี่ยนเป็นบอทชั่วคราวทันทีและเปิด Modal เฉพาะเครื่องนั้น
            cp.isBot = true;
            cp.temporaryBot = true;
            cp.returnRequested = false;
            cp.nextTurnTimeLimit = 40;

            broadcastAnnounce(`🤖 [PID:${cp.id}] หมดเวลาครบ 3 ครั้ง ระบบเปลี่ยนเป็นบอทชั่วคราวเพื่อให้เกมดำเนินต่อไป`);
            syncStateToAll();
            startHostBotResponseTimer(cp.id);

            // ให้บอทควบคุมตัวละครนั้นต่อทันทีในตานี้โดยไม่หยุดเกม
            checkBotTurn();
        }
    } finally {
        isProcessingTurnTimeout = false;
    }
}

function syncStateToAll() {
    let activePlayers = gameState.players.filter(p => !p.isOut);
    if(activePlayers.length <= 1) return;

    const currentTopDiscard = discardPile[discardPile.length - 1] || null;
    gameState.players.forEach(p => {
        if (p.hand && p.hand.length > 0) {
            sortPlayerHand(p.hand, currentTopDiscard, p.hasDrawnTurn, p.hasDiscardedTurn, p.hasFlowedThisTurn);
        }
    });

    let publicState = {
        status: gameState.status, turnId: gameState.players[gameState.turnIndex].id, topDiscard: currentTopDiscard,
        deckCount: deck.length, discardPileCount: discardPile.length, skipVotes: gameState.skipPreVotes || [],
        playersInfo: gameState.players.map(p => ({
            id: p.id,
            name: p.name,
            colorHex: p.colorInfo ? p.colorInfo.hex : 'var(--red)',
            colorName: p.colorInfo ? p.colorInfo.name : 'แดง',
            isBot: p.isBot,
            temporaryBot: !!p.temporaryBot,
            returnRequested: !!p.returnRequested,
            permanentBot: !!p.permanentBot,
            points: p.points,
            cardCount: p.hand.length,
            isOut: p.isOut
        }))
    };

    gameState.players.forEach(p => {
        // ซิงก์ข้อมูลให้ผู้เล่นจริงทุกคน รวมถึงผู้เล่นที่ถูกเปลี่ยนเป็นบอทชั่วคราว (ตราบใดที่ยังไม่ได้ออกจากห้องถาวร)
        const isHumanSlot = !p.originalIsBot && !String(p.id).startsWith('bot_');
        if (isHumanSlot && !p.permanentBot) {
            let privateState = {
                hand: p.hand,
                hasDrawn: p.hasDrawnTurn,
                hasDiscarded: p.hasDiscardedTurn,
                discardedRank: p.discardedRankThisTurn,
                points: p.points,
                turnCount: p.turnCount,
                hasFinishedFirstTurn: p.hasFinishedFirstTurn,
                hasFlowedThisTurn: p.hasFlowedThisTurn,
                temporaryBot: !!p.temporaryBot,
                returnRequested: !!p.returnRequested
            };
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
// [พื้นที่สำหรับพัฒนาต่อ #2] ระบบ Animation (เพิ่มมิติการแสดงผล Knock / Special Win)
// =========================================================
window.playKangAnimation = (callerId, winType) => {
    const animScreen = document.getElementById('kang-animation-screen');
    const text1 = document.getElementById('kang-anim-text1');
    const text2 = document.getElementById('kang-anim-text2');
    const avatar = document.getElementById('kang-anim-avatar');

    animScreen.style.display = 'flex';
    
    // --- [เพิ่มมิติการแสดงผล] Visual Trigger: Screen Flash ---
    animScreen.style.animation = 'none';
    void animScreen.offsetWidth; // trigger reflow
    animScreen.style.animation = 'flashScreen 0.4s ease-out';
    
    text1.style.opacity = '0';
    text2.style.opacity = '0';
    
    let pInfo = globalPlayersMap[callerId];
    let pColorHex = pInfo ? pInfo.colorHex : 'white';
    
    avatar.style.color = pColorHex;
    avatar.className = 'kang-anim-active';
    
    // --- [เพิ่มมิติการแสดงผล] Visual Trigger: Dramatic Avatar Pop ---
    avatar.style.transform = 'scale(0.1) rotate(-15deg)';
    avatar.style.transition = 'transform 0.6s cubic-bezier(0.34, 1.56, 0.64, 1)';
    setTimeout(() => avatar.style.transform = 'scale(1) rotate(0deg)', 50);

    // --- [เพิ่มมิติการแสดงผล] Visual Trigger: Screen Shake ---
    if (winType === 'cang25' || winType === 'knock' || winType === 'special') {
        document.body.style.animation = 'shake 0.5s ease-in-out';
        setTimeout(() => document.body.style.animation = '', 500);
    }

    let soundKey = '';
    if (winType === 'cang25') soundKey = 'cang25';
    else if (winType === 'knock') soundKey = 'knock';
    else if (winType === 'special') soundKey = 'special';
    else soundKey = 'cang';

    playSound(soundKey);

    let callerName = resolveName(callerId);

    setTimeout(() => {
        if (winType === 'cang') {
            text1.innerText = t('kangAnim1_cang', callerName);
        } else if (winType === 'cang25') {
            text1.innerText = t('kangAnim1_cang25', callerName);
        } else if (winType === 'knock') {
            text1.innerText = t('kangAnim1_knock');
        } else if (winType === 'special') {
            text1.innerText = t('kangAnim1_special');
        }
        text1.style.opacity = '1'; announce(text1.innerText, false);
    }, 1200);

    setTimeout(() => {
        if (winType === 'cang' || winType === 'cang25') {
            text2.innerText = t('kangAnim2_cang');
        } else if (winType === 'knock') {
            text2.innerText = t('kangAnim2_knock', callerName);
        } else if (winType === 'special') {
            text2.innerText = t('kangAnim2_special', callerName);
        }
        text2.style.opacity = '1'; announce(text2.innerText, false);
    }, 3000);

    setTimeout(() => {
        animScreen.style.display = 'none'; avatar.className = '';
    }, 4500);
};

function broadcastKangAnimation(callerId, winType) {
    clearAutoEndTurnTimer();
    gameState.status = 'TRANSITION'; clearInterval(turnTimerInterval); syncStateToAll();
    let data = { type: 'KANG_ANIMATION', callerId, winType };
    if (myPeerId) window.playKangAnimation(callerId, winType);
    guestConnections.forEach(c => {
        if (c && c.peer !== myPeerId) {
            sendSoundMessageToConn(c, data);
        }
    });
}
// =========================================================

function nextTurn() {
    clearAutoEndTurnTimer();
    isEndingTurn = false;
    activeBotSequencePlayerId = null;

    let prevPlayer = gameState.players[gameState.turnIndex];
    prevPlayer.hasDrawnTurn = false; prevPlayer.hasDiscardedTurn = false; prevPlayer.hasFlowedThisTurn = false; prevPlayer.discardedRankThisTurn = null;

    do { gameState.turnIndex = (gameState.turnIndex + 1) % gameState.players.length;
    } while (gameState.players[gameState.turnIndex].isOut || gameState.players[gameState.turnIndex].hand.length === 0);
    
    let currentTurnPlayer = gameState.players[gameState.turnIndex];
    if (restorePlayerIfReturnRequested(currentTurnPlayer)) {
        broadcastAnnounce(`✅ [PID:${currentTurnPlayer.id}] กลับมาควบคุมการเล่นตามปกติแล้ว`);
    }

    triggerSound('turn'); syncStateToAll(); startTurnTimer(); checkBotTurn();
}

function checkBotTurn() {
    let cp = gameState.players[gameState.turnIndex];
    if (!cp || !cp.isBot || gameState.status !== 'PLAYING') return;
    if (activeBotSequencePlayerId === cp.id) return;
    activeBotSequencePlayerId = cp.id;
    setTimeout(() => playBotSequence(cp), 1000);
}

async function playBotSequence(cp) {
    try {
        await delay(2000);
        if (gameState.status !== 'PLAYING' || gameState.players[gameState.turnIndex]?.id !== cp.id || !cp.isBot) return;

        let handVal = cp.hand.reduce((sum, c) => sum + c.val, 0);
        if (handVal <= 5 && discardPile.length > 0 && !cp.hasDrawnTurn && !cp.hasDiscardedTurn && !cp.hasFlowedThisTurn) {
            processPlayerAction(cp.id, 'KANG', 0, true);
            return;
        }

        let topCard = discardPile[discardPile.length - 1];
        if (topCard && !cp.hasDrawnTurn && !cp.hasDiscardedTurn) {
            let flowIndex = cp.hand.findIndex(c => c.rank === topCard.rank);
            if (flowIndex !== -1) {
                processPlayerAction(cp.id, 'FLOW', flowIndex, true);
                await delay(2500);
                while (gameState.status === 'PLAYING' && gameState.players[gameState.turnIndex]?.id === cp.id && cp.isBot) {
                    let matchIdx = cp.hand.findIndex(c => c.rank === topCard.rank);
                    if (matchIdx !== -1) {
                        processPlayerAction(cp.id, 'FLOW', matchIdx, true);
                        await delay(1500);
                    } else break;
                }
                await delay(2000);
                if (gameState.status === 'PLAYING' && gameState.players[gameState.turnIndex]?.id === cp.id && cp.isBot) {
                    processPlayerAction(cp.id, 'END_TURN', 0, true);
                }
                return;
            }
        }

        if (!cp.hasDrawnTurn && !cp.hasFlowedThisTurn) {
            processPlayerAction(cp.id, 'DRAW', 0, true);
            await delay(2000);
        }
        if (gameState.status !== 'PLAYING' || gameState.players[gameState.turnIndex]?.id !== cp.id || !cp.isBot) return;

        if (!cp.hasDiscardedTurn && !cp.hasFlowedThisTurn && cp.hand.length > 0) {
            let rankCounts = {};
            cp.hand.forEach(c => rankCounts[c.rank] = (rankCounts[c.rank] || 0) + 1);
            let bestRank = null, maxCount = 0, maxVal = -1;
            for (let r in rankCounts) {
                let count = rankCounts[r];
                let val = getCardValue(r);
                if (count > maxCount || (count === maxCount && val > maxVal)) {
                    maxCount = count;
                    maxVal = val;
                    bestRank = r;
                }
            }
            while (gameState.status === 'PLAYING' && gameState.players[gameState.turnIndex]?.id === cp.id && cp.isBot) {
                let matchIdx = cp.hand.findIndex(c => c.rank === bestRank);
                if (matchIdx !== -1) {
                    processPlayerAction(cp.id, 'DISCARD', matchIdx, true);
                    await delay(1500);
                } else break;
            }
        }

        await delay(2000);
        if (gameState.status === 'PLAYING' && gameState.players[gameState.turnIndex]?.id === cp.id && cp.isBot) {
            processPlayerAction(cp.id, 'END_TURN', 0, true);
        }
    } finally {
        if (activeBotSequencePlayerId === cp.id) {
            activeBotSequencePlayerId = null;
        }
    }
}

function resolveSpecialWin(callerId, winType) {
    // ใช้งาน Special Win Animation แทรกก่อนทำการสรุปผลและคำนวณแต้ม
    broadcastKangAnimation(callerId, 'special');
    
    setTimeout(() => {
        window.nextRoundHostId = callerId;
        gameState.status = 'TRANSITION'; 
        clearInterval(turnTimerInterval);
        
        let active = gameState.players.filter(p => !p.isOut);
        let beforePoints = {};
        active.forEach(p => beforePoints[p.id] = p.points);
        
        let caller = active.find(p => p.id === callerId);
        let others = active.filter(p => p.id !== callerId);
        
        let aceCount = caller.hand.filter(c => c.rank === 'A').length;
        let payPerPerson = 0;
        
        if (winType === '3A') { payPerPerson = 40; } else { payPerPerson = 20 + (aceCount * 10); }
        
        others.forEach(p => { p.points -= payPerPerson; if(p.points < 0) p.points = 0; caller.points += payPerPerson; });
        
        let title = `[PID:${caller.id}] ชนะพิเศษกติกา: ${winType}!`;
        let winners = [caller.id]; let losers = [];
        let cardDetailLines = [`[PID:${caller.id}] เปิดไพ่ชนะกติกา ${winType}`];
        let coinDetailLines = [];
        
        active.forEach(p => {
            let diff = p.points - beforePoints[p.id];
            if (diff > 0) { coinDetailLines.push(`[PID:${p.id}] ได้รับ ${diff} เหรียญ`); } 
            else if (diff < 0) { losers.push(p.id); coinDetailLines.push(`[PID:${p.id}] เสีย ${Math.abs(diff)} เหรียญ`); } 
            else { coinDetailLines.push(`[PID:${p.id}] ไม่ได้ไม่เสียเหรียญ`); }
        });
        
        let resultDetail = cardDetailLines.join('<br>') + '<br><br>' + coinDetailLines.join('<br>');
        syncStateToAll(); broadcastTransition('จบเกมด้วยกติกาพิเศษ! กำลังสรุปผล...', 4000);
        
        setTimeout(() => {
            gameState.status = 'END'; syncStateToAll();
            let isFinalEnd = gameState.players.some(p => p.points === 0);
            let modalData = { type: 'SHOW_RESULT', title, detail: resultDetail, isEnd: isFinalEnd, winners, losers };
            if(myPeerId) window.showResultModal(title, resultDetail, isFinalEnd, winners, losers);
            guestConnections.forEach(c => {
                if (c && c.peer !== myPeerId) {
                    sendSoundMessageToConn(c, modalData);
                }
            });
        }, 4000);
    }, 4500);
}

function resetPlayerTimeoutOnValidHumanAction(p) {
    if (!p || p.isBot || p.temporaryBot || p.permanentBot) return;
    p.consecutiveTimeouts = 0;
    p.nextTurnTimeLimit = 40;
}

function processPlayerAction(pId, action, cardIndex, isInternalBotOrSystem = false) {
    let p = gameState.players.find(x => x.id === pId);
    if (!p || p.isOut) return;

    // ป้องกันผู้เล่นจริงส่งคำสั่งแทรกในขณะที่ตัวละครถูกควบคุมโดยบอท (ชั่วคราวหรือถาวร)
    if (!isInternalBotOrSystem && (p.isBot || p.temporaryBot || p.permanentBot)) {
        return;
    }

    if (action === 'SPECIAL_WIN') {
        if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(p);
        resolveSpecialWin(p.id, cardIndex);
        return;
    }

    if (gameState.status === 'PRE_GAME') {
        if (action === 'SKIP_PRE' && !p.isBot) {
            if (!gameState.skipPreVotes) gameState.skipPreVotes = [];
            if (!gameState.skipPreVotes.includes(pId)) {
                gameState.skipPreVotes.push(pId); syncStateToAll();
                let realPlayers = gameState.players.filter(x => !x.isOut && !x.isBot);
                if (gameState.skipPreVotes.length >= realPlayers.length) {
                    clearInterval(turnTimerInterval); triggerSound('60'); gameState.status = 'PLAYING';
                    let firstPlayer = gameState.players[gameState.turnIndex] || gameState.players[0];
                    restorePlayerIfReturnRequested(firstPlayer);
                    broadcastAnnounce(`📢 ผู้เล่นทุกคนพร้อมแล้ว! ขอเชิญ [PID:${firstPlayer.id}] จั่วไพ่ใบแรกและทิ้งลงกองเพื่อเริ่มเกมได้เลยครับ`);
                    syncStateToAll(); startTurnTimer(); checkBotTurn();
                }
            }
            return;
        }

        let handSum = p.hand.reduce((sum, c) => sum + c.val, 0);
        let isFirstPlayer = (pId === gameState.players[0].id);

        if (action === 'KANG_25' && handSum >= 25) {
            broadcastKangAnimation(p.id, 'cang25');
            setTimeout(() => resolveKang(p.id, false, false, true), 4500);
        } else if (action === 'KANG' && isFirstPlayer && handSum < 25) {
            broadcastKangAnimation(p.id, 'cang');
            setTimeout(() => resolveKang(p.id, false, false, false), 4500);
        }
        return;
    }

    if (action === 'FLOW') {
        if (gameState.status !== 'PLAYING') return;
        let flowPlayer = gameState.players.find(x => x.id === pId);
        if (!flowPlayer || flowPlayer.isOut) return;
        
        // --- การตั้งค่าป้องกันที่ 1 ---
        // ป้องกันการไหลถ้าผู้เล่นคนนั้นทำการจั่วไพ่แล้ว หรือทิ้งไพ่ปกติไปแล้วในเทิร์นตัวเอง
        if (flowPlayer.hasDrawnTurn || flowPlayer.hasDiscardedTurn) {
            return;
        }

        // --- เพิ่ม Logic Guard ป้องกันการไหลแซง Turn ---
        let currentPlayerId = gameState.players[gameState.turnIndex].id;
        if (pId !== currentPlayerId) {
            return; // ปฏิเสธ Action หากไม่ใช่ Turn ของผู้เล่นที่กดไหล
        }
        // ----------------------------------------------

        let topCard = discardPile[discardPile.length - 1]; let card = flowPlayer.hand[cardIndex];
        if (topCard && card && card.rank === topCard.rank) {
            if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(flowPlayer);
            clearAutoEndTurnTimer();
            triggerSound('follow'); let dropped = flowPlayer.hand.splice(cardIndex, 1)[0]; discardPile.push(dropped);
            flowPlayer.hasFlowedThisTurn = true; 

            let victimId = gameState.flowSourceId;
            let victim = gameState.players.find(v => v.id === victimId);
            if(!victim || victim.isOut) {
                let victimIndex = (gameState.turnIndex - 1 + gameState.players.length) % gameState.players.length;
                while(gameState.players[victimIndex].isOut) victimIndex = (victimIndex - 1 + gameState.players.length) % gameState.players.length;
                victim = gameState.players[victimIndex];
            }

            if (victim) {
                let flowMult = dropped.rank === 'A' ? 2 : 1; let penalty = 10 * flowMult;
                victim.points -= penalty; if(victim.points < 0) victim.points = 0; flowPlayer.points += penalty;
                broadcastAnnounce(`[PID:${flowPlayer.id}] ไหลไพ่ ${THAI_RANKS[dropped.rank]} ${THAI_SUITS[dropped.suit]}! ได้รับ ${penalty} เหรียญ จาก [PID:${victim.id}]`);
                broadcastFloatSC(victim.id, -penalty); broadcastFloatSC(flowPlayer.id, penalty);
            }

            gameState.topCardOwnerId = flowPlayer.id; 

            if(flowPlayer.hand.length === 0) { 
                broadcastKangAnimation(flowPlayer.id, 'knock');
                setTimeout(() => resolveKang(flowPlayer.id, false, 'FLOW_KNOCK'), 4500); 
            } 
            else {
                const remainingFlowable = flowPlayer.hand.filter(c => c.rank === topCard.rank).length;
                const shouldAutoEndAfterFlow = (!flowPlayer.isBot && !isInternalBotOrSystem && remainingFlowable === 0);

                gameState.status = 'TRANSITION'; clearInterval(turnTimerInterval); syncStateToAll();
                broadcastTransition('ไหลไพ่สำเร็จ! รอสักครู่...', 2500);
                setTimeout(() => {
                    gameState.status = 'PLAYING';
                    syncStateToAll();
                    startTurnTimer();
                    if (flowPlayer.isBot) {
                        checkBotTurn();
                    } else if (shouldAutoEndAfterFlow) {
                        scheduleAutoEndTurn(flowPlayer.id);
                    }
                }, 2500);
            }
        }
        return;
    }

    let cp = gameState.players[gameState.turnIndex];
    if(cp.id !== pId || gameState.status !== 'PLAYING') return; 

    if(action === 'DRAW' && !cp.hasDrawnTurn && !cp.hasFlowedThisTurn) {
        if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(cp);
        if(deck.length > 0) {
            triggerSound('jua'); let drawnCard = deck.pop();
            cp.hand.push(drawnCard); cp.hasDrawnTurn = true;
            let publicMsg = `[PID:${cp.id}] จั่วไพ่ 1 ใบ`;
            if (cp.isBot && !cp.temporaryBot) { broadcastAnnounce(publicMsg); } 
            else {
                let privateMsg = `[PID:${cp.id}] จั่วได้ไพ่ ${THAI_RANKS[drawnCard.rank]} ${THAI_SUITS[drawnCard.suit]}`;
                if (myPeerId === cp.id) announce(privateMsg); else announce(publicMsg);
                guestConnections.forEach(c => {
                    if (c.peer === cp.id) c.send({ type: 'ANNOUNCE', msg: privateMsg }); else c.send({ type: 'ANNOUNCE', msg: publicMsg });
                });
            }
            syncStateToAll();
        } else resolveKang(cp.id, true); // กองไพ่หมดให้จบเกมทันที
    }
    else if(action === 'DISCARD') {
        let card = cp.hand[cardIndex];
        if (!card) return;
        if (!cp.hasDiscardedTurn && cp.hasDrawnTurn) {
            if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(cp);
            clearAutoEndTurnTimer();
            triggerSound('select'); let dropped = cp.hand.splice(cardIndex, 1)[0]; discardPile.push(dropped);
            
            let isFirstDiscardInTurn = !cp.hasDiscardedTurn;
            cp.hasDiscardedTurn = true; cp.discardedRankThisTurn = dropped.rank;
            
            if (isFirstDiscardInTurn) {
                gameState.flowSourceId = cp.id;
            }

            gameState.topCardOwnerId = cp.id;
            broadcastAnnounce(`[PID:${cp.id}] วางไพ่ ${THAI_RANKS[dropped.rank]} ${THAI_SUITS[dropped.suit]}`);
            if(cp.hand.length === 0) {
                broadcastKangAnimation(cp.id, 'knock');
                setTimeout(() => resolveKang(cp.id, false, 'DRAW_KNOCK'), 4500);
            } else {
                syncStateToAll();
                if (!cp.isBot && !isInternalBotOrSystem) {
                    const remainingSameRank = cp.hand.filter(c => c.rank === dropped.rank).length;
                    if (remainingSameRank === 0) {
                        scheduleAutoEndTurn(cp.id);
                    }
                }
            }
        } 
        else if (cp.hasDiscardedTurn && card.rank === cp.discardedRankThisTurn) {
            if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(cp);
            clearAutoEndTurnTimer();
            triggerSound('select'); let dropped = cp.hand.splice(cardIndex, 1)[0]; discardPile.push(dropped);
            broadcastAnnounce(`[PID:${cp.id}] วางไพ่ ${THAI_RANKS[dropped.rank]} ${THAI_SUITS[dropped.suit]} เพิ่มเติม (เลขเดียวกัน)`);
            if(cp.hand.length === 0) {
                broadcastKangAnimation(cp.id, 'knock');
                setTimeout(() => resolveKang(cp.id, false, 'DRAW_KNOCK'), 4500);
            } else {
                syncStateToAll();
                if (!cp.isBot && !isInternalBotOrSystem) {
                    const remainingSameRank = cp.hand.filter(c => c.rank === dropped.rank).length;
                    if (remainingSameRank === 0) {
                        scheduleAutoEndTurn(cp.id);
                    }
                }
            }
        }
    }
    else if(action === 'END_TURN' && (cp.hasDiscardedTurn || cp.hasFlowedThisTurn)) {
        if (isEndingTurn) return;
        isEndingTurn = true;
        clearAutoEndTurnTimer();
        if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(cp);
        cp.hasFinishedFirstTurn = true;
        cp.turnCount = (cp.turnCount || 0) + 1;
        triggerSound('select'); broadcastAnnounce(`[PID:${cp.id}] กดยืนยันจบตา`);
        gameState.status = 'TRANSITION'; clearInterval(turnTimerInterval); syncStateToAll();
        broadcastTransition('กำลังเปลี่ยนตาถัดไป...', 2000);
        setTimeout(() => { gameState.status = 'PLAYING'; nextTurn(); }, 2000);
    }
    else if(action === 'KANG' && !cp.hasDrawnTurn && !cp.hasFlowedThisTurn && discardPile.length > 0) {
        if (!isInternalBotOrSystem) resetPlayerTimeoutOnValidHumanAction(cp);
        broadcastKangAnimation(cp.id, 'cang'); setTimeout(() => resolveKang(cp.id), 4500);
    }
}

function resolveKang(callerId, isDeckEmpty = false, winReason = false, isKang25 = false) {
    clearAutoEndTurnTimer();
    window.nextRoundHostId = callerId;
    gameState.status = 'TRANSITION'; clearInterval(turnTimerInterval);
    let active = gameState.players.filter(p => !p.isOut);
    let totalPlayersInRoom = active.length;
    let beforePoints = {}; active.forEach(p => beforePoints[p.id] = p.points);

    active.forEach(p => { 
        p.handSum = p.hand.reduce((s, c) => s + c.val, 0); 
        p.aceCount = p.hand.filter(c => c.rank === 'A').length;
    });
    
    let caller = active.find(p => p.id === callerId); 
    let resultDetail = "", title = "";
    const getBasePay = (player) => 10 + (player.aceCount * 10);

    if (winReason === 'FLOW_KNOCK' || winReason === 'DRAW_KNOCK' || winReason === true) {
        let basePay = (winReason === 'DRAW_KNOCK') ? 20 : 10;
        let reasonStr = (winReason === 'DRAW_KNOCK') ? 'จั่วแล้วทิ้งไพ่หมดมือ' : 'ไหลไพ่หมดมือน็อก';
        if (winReason === true) reasonStr = 'น็อค!';
        
        title = `[PID:${caller.id}] ${reasonStr}! (ชนะรับ ${basePay} เหรียญ จากทุกคน)`;
        active.forEach(p => { 
            if(p.id !== caller.id) { p.points -= basePay; if(p.points < 0) p.points = 0; caller.points += basePay; } 
        });
    } else if (isDeckEmpty) {
        let lowest = active.reduce((min, p) => p.handSum < min.handSum ? p : min, active[0]);
        let basePay = getBasePay(lowest);
        title = `กองไพ่หมด! [PID:${lowest.id}] แต้มต่ำสุด ${lowest.handSum} แต้ม (ชนะ)`;
        active.forEach(p => { 
            if(p.id !== lowest.id) { p.points -= basePay; if(p.points < 0) p.points = 0; lowest.points += basePay; } 
        });
    } else { 
        let others = active.filter(p => p.id !== caller.id);
        let lowestOther = others.reduce((min, p) => p.handSum < min.handSum ? p : min, others[0]);

        if (caller.handSum <= lowestOther.handSum) { // เปลี่ยนจาก < เป็น <= เพื่อให้ชนะเสมอเมื่อแต้มต่ำสุดเท่ากัน
            let payAmount = getBasePay(caller); if (isKang25) payAmount = payAmount * 2; 
            title = `[PID:${caller.id}] ${isKang25?'แคง 25':'แคง'} สำเร็จ!`;
            others.forEach(p => { p.points -= payAmount; if(p.points < 0) p.points = 0; caller.points += payAmount; });
        } else {
            let payAmount = getBasePay(lowestOther) * totalPlayersInRoom; 
            title = `[PID:${caller.id}] ${isKang25?'แคง 25':'แคง'} แหก! ([PID:${lowestOther.id}] ชนะ)`;
            caller.points -= payAmount; if(caller.points < 0) caller.points = 0; lowestOther.points += payAmount;
        }
    }

    let winners = []; let losers = [];
    let cardDetailLines = []; let coinDetailLines = [];

    active.forEach(p => {
        let diff = p.points - beforePoints[p.id];
        if (diff > 0) { winners.push(p.id); coinDetailLines.push(`[PID:${p.id}] ได้รับ ${diff} เหรียญ`); } 
        else if (diff < 0) { losers.push(p.id); coinDetailLines.push(`[PID:${p.id}] เสีย ${Math.abs(diff)} เหรียญ`); } 
        else { coinDetailLines.push(`[PID:${p.id}] ไม่ได้ไม่เสียเหรียญ`); }
        cardDetailLines.push(`[PID:${p.id}] เปิดไพ่มาแล้วได้แต้มรวม ${p.handSum} แต้ม`);
    });

    resultDetail = cardDetailLines.join('<br>') + '<br><br>' + coinDetailLines.join('<br>');
    syncStateToAll(); broadcastTransition('จบเกม! กำลังสรุปผล...', 4000);

    setTimeout(() => {
        gameState.status = 'END'; syncStateToAll();
        let isFinalEnd = gameState.players.some(p => p.points === 0);
        let modalData = { type: 'SHOW_RESULT', title, detail: resultDetail, isEnd: isFinalEnd, winners, losers };
        if(myPeerId) window.showResultModal(title, resultDetail, isFinalEnd, winners, losers);
        guestConnections.forEach(c => {
            if (c && c.peer !== myPeerId) {
                sendSoundMessageToConn(c, modalData);
            }
        });
    }, 4000);
}

window.closeResultModal = () => {
    if (window.isStartingRound) return;
    window.isStartingRound = true;
    let btn = document.getElementById('btn-next-round');
    if (btn) { btn.disabled = true; btn.style.display = 'none'; }
    document.getElementById('result-modal').style.display = 'none';
    updateHomeBtnVisibility();
    if(isHost) { guestConnections.forEach(c => sendSoundMessageToConn(c, { type: 'START_COUNTDOWN' })); doCountdownAndStart(); }
};

// =========================================================
// Client Rendering Logic
// =========================================================
function renderClientGame(publicState, privateState) {
    window.currentPublicState = publicState; // [แคช state ล่าสุดไว้ใช้ดึงเวลาเสกไพ่ใน Client]

    if (publicState.status === 'PRE_GAME' || publicState.status === 'PLAYING') { document.getElementById('result-modal').style.display = 'none'; }
    updateHomeBtnVisibility();

    localPlayerState = privateState;
    let topC = publicState.topDiscard;
    if (localPlayerState && Array.isArray(localPlayerState.hand)) {
        sortPlayerHand(
            localPlayerState.hand,
            topC,
            localPlayerState.hasDrawn,
            localPlayerState.hasDiscarded,
            localPlayerState.hasFlowedThisTurn
        );
    }

    if (publicState.status === 'PRE_GAME') {
        if (window._pendingPreGameBaseMsg && !window._announcedPreGameRound && localPlayerState.hand && localPlayerState.hand.length > 0) {
            const baseMsg = window._pendingPreGameBaseMsg;
            window._pendingPreGameBaseMsg = null;
            announce(baseMsg);
        }
    } else {
        window._pendingPreGameBaseMsg = null;
        window._announcedPreGameRound = false;
    }

    const turnId = publicState.turnId;
    const isFirstPlayer = (myPeerId === publicState.playersInfo[0].id);
    const myPublicInfo = publicState.playersInfo.find(p => p.id === myPeerId);
    const isControlledByBot = !!(localPlayerState.temporaryBot || (myPublicInfo && (myPublicInfo.isBot || myPublicInfo.temporaryBot)));
    const isMyTurn = (turnId === myPeerId) && !isControlledByBot;
    
    publicState.playersInfo.forEach(p => {
        let pName = (p.name && String(p.name).trim() !== '') ? String(p.name).trim() : (globalPlayersMap[p.id] && globalPlayersMap[p.id].name ? globalPlayersMap[p.id].name : 'ผู้เล่น');
        globalPlayersMap[p.id] = { isBot: p.isBot, temporaryBot: p.temporaryBot, permanentBot: p.permanentBot, name: pName, colorName: p.colorName, colorHex: p.colorHex };
    });
    let currPlayerDisp = resolveName(turnId);
    
    const turnInd = document.getElementById('turn-indicator');
    const timerBarContainer = document.getElementById('timer-bar-container');
    
    if (publicState.status === 'END') {
        if(timerBarContainer) timerBarContainer.style.display = 'none';
        turnInd.style.display = 'none'; 
        document.getElementById('main-table-area').style.display = 'none';
        document.getElementById('action-bar-container').style.display = 'none';
    } else {
        turnInd.style.display = 'block'; 
        if(timerBarContainer) timerBarContainer.style.display = 'block';
        document.getElementById('main-table-area').style.display = 'flex';
        document.getElementById('action-bar-container').style.display = 'flex';
        if (publicState.status === 'PRE_GAME') {
            turnInd.innerText = t('preGameTurnInd');
        } else {
            let oldTurnTxt = turnInd.innerText;
            let newTurnTxt = (turnId === myPeerId && !isControlledByBot) ? t('yourTurnInd', currPlayerDisp) : t('otherTurnInd', currPlayerDisp);
            if (oldTurnTxt !== newTurnTxt && publicState.status === 'PLAYING') {
                turnInd.innerText = newTurnTxt;
                let topTxt = topC ? formatCardName(topC.rank, topC.suit) : t('topCardEmpty');
                if (turnId === myPeerId && !isControlledByBot) {
                    let playableCount = (topC && !localPlayerState.hasDrawn && !localPlayerState.hasFlowedThisTurn)
                        ? (localPlayerState.hand || []).filter(c => c.rank === topC.rank).length
                        : 0;
                    let playHintTxt = playableCount > 0 ? t('playHintCan', playableCount) : t('playHintCant');
                    announce(t('yourTurnAnnounce', currPlayerDisp, topTxt, playHintTxt));
                    if (playableCount === 0 && !localPlayerState.hasDrawn && !localPlayerState.hasFlowedThisTurn) {
                        scheduleFocusOnDrawButton();
                    }
                } else {
                    announce(t('otherTurnAnnounce', currPlayerDisp, topTxt), false);
                }
            }
        }
    }

    const renderStatusBar = (containerId) => {
        const container = document.getElementById(containerId);
        if (!container) return;
        container.innerHTML = '';
        let groupAriaLabel = "";

        publicState.playersInfo.forEach((p) => {
            if (!p.isOut) {
                let dispName = resolveName(p.id);
                groupAriaLabel += t('statusPlayerAria', dispName, p.points);
                let item = document.createElement('div');
                item.className = `status-avatar-item ${p.id === turnId && publicState.status === 'PLAYING' ? 'active-turn' : ''}`;
                item.id = `avatar-${containerId}-${p.id}`;
                item.style.borderColor = p.colorHex;
                item.innerHTML = `
                    <div class="status-coin-header" aria-hidden="true">${t('statusCoinsText', p.points)}</div>
                    <div class="status-avatar-icon" style="background-color:${p.colorHex}; color:#fff;" aria-hidden="true">👤</div>
                    <div class="status-player-label" style="color:${p.colorHex};" aria-hidden="true">${dispName}</div>
                `;
                container.appendChild(item);
            }
        });
        container.setAttribute('tabindex', '0');
        container.setAttribute('role', 'group');
        container.setAttribute('aria-label', t('statusGroupAria', groupAriaLabel));
    };

    renderStatusBar('players-status-bar'); renderStatusBar('result-status-bar');

    const btnDraw = document.getElementById('btn-draw'); const btnKang = document.getElementById('btn-kang');
    const btnKang25 = document.getElementById('btn-kang25'); const btnEnd = document.getElementById('btn-end-turn');
    const btnSkip = document.getElementById('btn-skip-pre'); const btnSpecial = document.getElementById('btn-special');

    if (topC) currentTopCardText = t('topCardCurrent', formatCardName(topC.rank, topC.suit));
    else currentTopCardText = t('topCardEmptyFull');

    let hasVoted = (publicState.skipVotes && publicState.skipVotes.includes(myPeerId)) || isControlledByBot;

    if (publicState.status === 'PRE_GAME') {
        let myHandSum = (localPlayerState.hand || []).reduce((s, c) => s + (c.val || getCardValue(c.rank)), 0);
        btnDraw.style.display = 'none'; btnEnd.style.display = 'none';
        
        if (isFirstPlayer) {
            if (myHandSum < 25) {
                btnKang.style.display = 'inline-block'; btnKang.disabled = hasVoted; 
                btnKang25.style.display = 'none'; btnKang25.disabled = true;
            } else {
                btnKang.style.display = 'none'; btnKang.disabled = true;
                btnKang25.style.display = 'inline-block'; btnKang25.disabled = hasVoted; 
            }
        } else {
            btnKang.style.display = 'none'; btnKang.disabled = true;
            if (myHandSum >= 25) { btnKang25.style.display = 'inline-block'; btnKang25.disabled = hasVoted; } 
            else { btnKang25.style.display = 'none'; btnKang25.disabled = true; }
        }
        btnSkip.style.display = 'inline-block'; btnSkip.disabled = hasVoted;

    } else if (publicState.status === 'PLAYING' || publicState.status === 'TRANSITION') {
        btnKang25.style.display = 'none'; btnSkip.style.display = 'none';
        btnDraw.style.display = 'inline-block'; btnEnd.style.display = 'inline-block'; btnKang.style.display = 'inline-block';

        if(isMyTurn && publicState.status === 'PLAYING' && !isControlledByBot) {
            btnDraw.disabled = localPlayerState.hasDrawn || localPlayerState.hasFlowedThisTurn;
            btnKang.disabled = localPlayerState.hasDrawn || localPlayerState.hasFlowedThisTurn || publicState.discardPileCount === 0;
            btnEnd.disabled = !(localPlayerState.hasDiscarded || localPlayerState.hasFlowedThisTurn);
        } else { btnDraw.disabled = true; btnKang.disabled = true; btnEnd.disabled = true; }
    }

    currentSpecialWinType = getSpecialWinType(localPlayerState.hand, localPlayerState);

    if (currentSpecialWinType && (publicState.status === 'PRE_GAME' || publicState.status === 'PLAYING') && !isControlledByBot) {
        btnSpecial.style.display = 'inline-block';
        btnSpecial.innerText = getSpecialWinDisplayName(currentSpecialWinType);
        btnSpecial.disabled = false;
    } else {
        btnSpecial.style.display = 'none';
        btnSpecial.disabled = true;
    }

    const deckCountTxt = document.getElementById('deck-count-text');
    if (deckCountTxt) {
        deckCountTxt.innerText = publicState.deckCount;
    }

    const deckAccessibleObject = document.getElementById('deck-accessible-object');
    if (deckAccessibleObject) {
        deckAccessibleObject.setAttribute(
            'aria-label',
            t('deckAria', publicState.deckCount)
        );
    }

    let dp = document.getElementById('discard-pile');
    
    // --- [เพิ่มมิติการแสดงผล] Visual Trigger: อัปเดตกองทิ้ง ---
    if(topC) {
        let currentTopId = `${topC.rank}-${topC.suit}`;
        if (window._lastTopCardId !== currentTopId) {
            window._lastTopCardId = currentTopId;
            dp.style.animation = 'none';
            void dp.offsetWidth; // trigger reflow
            dp.style.animation = 'cardDropScale 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275)';
        }

        let cardNameStr = formatCardName(topC.rank, topC.suit);
        dp.innerHTML = `<div class="card ${['♥','♦'].includes(topC.suit)?'red':'black'}" aria-label="${t('discardPileTopAria', cardNameStr)}" role="img">
                           <div aria-hidden="true">${topC.rank}</div><div class="card-suit" aria-hidden="true">${topC.suit}</div>
                        </div>`;
    } else {
        window._lastTopCardId = null;
        dp.innerHTML = `<div style="color:#a0a0a0; font-size:0.85rem;">${t('discardPileEmptyHtml')}</div>`; 
        dp.setAttribute('aria-label', t('discardPileEmptyAria'));
    }

    const handUi = document.getElementById('my-hand-ui'); handUi.innerHTML = '';
    if (publicState.status !== 'END') {
        
        // --- การตั้งค่าป้องกันที่ 2 ---
        // ปรับเงื่อนไข canFlowCard ไม่ให้เกิดสถานะการไหลได้ ถ้าตัวผู้เล่นจั่วไพ่แล้ว หรือทิ้งไพ่รอบนี้ไปแล้ว
        (localPlayerState.hand || []).forEach((c, index) => {
            let cardEl = document.createElement('button');
            cardEl.className = `card ${['♥','♦'].includes(c.suit)?'red':'black'}`;
            
            // --- [เพิ่มมิติการแสดงผล] Visual Trigger: ไพ่ทยอยแจกเข้ามือ ---
            cardEl.style.animation = `cardSlideUp 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275) ${index * 0.08}s both`;

            let cardNameStr = formatCardName(c.rank, c.suit);
            let actionHint = ""; cardEl.disabled = true;
            
            let canFlowCard = (topC && c.rank === topC.rank && !localPlayerState.hasDrawn && !localPlayerState.hasDiscarded);

            if (publicState.status === 'PLAYING' && !isControlledByBot) {
                if (canFlowCard) {
                    if (isMyTurn) { // เป็น Turn ตัวเองเท่านั้นจึงจะไหลได้
                        cardEl.disabled = false; actionHint = t('hintFlow');
                        cardEl.style.boxShadow = "0 0 12px var(--blue)"; 
                        cardEl.onclick = () => window.clientAction('FLOW', index);
                    } else { // ไม่ใช่ Turn ตัวเอง ไม่สามารถกดได้ และทำให้แสดงสถานะ dim/disabled
                        cardEl.disabled = true;
                        cardEl.style.opacity = '0.5';
                        cardEl.style.cursor = 'not-allowed';
                    }
                }
                else if (isMyTurn) {
                    if (!localPlayerState.hasDiscarded) {
                        if (localPlayerState.hasDrawn && !localPlayerState.hasFlowedThisTurn) {
                            cardEl.disabled = false; actionHint = t('hintDiscard');
                            cardEl.onclick = () => window.clientAction('DISCARD', index);
                        }
                    } else if (c.rank === localPlayerState.discardedRank) {
                        cardEl.disabled = false; actionHint = t('hintDiscardMore');
                        cardEl.style.boxShadow = "0 0 12px var(--green)";
                        cardEl.onclick = () => window.clientAction('DISCARD', index);
                    }
                }
            }
            
            cardEl.setAttribute('aria-label', t('cardAria', cardNameStr, actionHint));
            cardEl.innerHTML = `<div aria-hidden="true">${c.rank}</div><div class="card-suit" aria-hidden="true">${c.suit}</div>`;
            handUi.appendChild(cardEl);
        });
    }
}

window.clientAction = (action, index = 0) => {
    const myPub = window.currentPublicState && window.currentPublicState.playersInfo
        ? window.currentPublicState.playersInfo.find(p => p.id === myPeerId)
        : null;
    if (localPlayerState.temporaryBot || (myPub && (myPub.isBot || myPub.temporaryBot))) {
        return;
    }

    // --- [เพิ่มมิติการแสดงผล] Visual Trigger: แอนิเมชันเมื่อผู้เล่นกดทำ Action ส่งข้อมูล ---
    try {
        if (action === 'DISCARD' || action === 'FLOW') {
            const handUi = document.getElementById('my-hand-ui');
            const cardEl = handUi ? handUi.children[index] : null;
            const discardPileEl = document.getElementById('discard-pile');
            
            if (cardEl && discardPileEl) {
                const cardRect = cardEl.getBoundingClientRect();
                const targetRect = discardPileEl.getBoundingClientRect();
                
                const clone = cardEl.cloneNode(true);
                clone.style.position = 'fixed';
                clone.style.left = `${cardRect.left}px`;
                clone.style.top = `${cardRect.top}px`;
                clone.style.width = `${cardRect.width}px`;
                clone.style.height = `${cardRect.height}px`;
                clone.style.zIndex = '9999';
                clone.style.transition = 'all 0.4s cubic-bezier(0.25, 1, 0.5, 1)';
                clone.style.pointerEvents = 'none';
                clone.style.margin = '0';
                document.body.appendChild(clone);
                
                cardEl.style.opacity = '0'; // ซ่อนไพ่บนมือทันทีเพื่อความสมูท
                
                setTimeout(() => {
                    clone.style.left = `${targetRect.left + (targetRect.width/2) - (cardRect.width/2)}px`;
                    clone.style.top = `${targetRect.top + (targetRect.height/2) - (cardRect.height/2)}px`;
                    clone.style.transform = `scale(0.8) rotate(${Math.random() * 30 - 15}deg)`;
                    clone.style.opacity = '0.7';
                }, 10);
                
                setTimeout(() => clone.remove(), 400);
            }
        } else if (action === 'DRAW') {
            const deckEl = document.getElementById('deck-group-box');
            if (deckEl) {
                const deckRect = deckEl.getBoundingClientRect();
                const clone = document.createElement('div');
                clone.className = 'card';
                // สร้างจำลองหลังไพ่คร่าวๆ สำหรับให้ลอยเข้ามือ
                clone.style.background = 'linear-gradient(135deg, var(--blue, #1e3c72), #2a5298)';
                clone.style.border = '2px solid #fff';
                clone.style.position = 'fixed';
                clone.style.left = `${deckRect.left}px`;
                clone.style.top = `${deckRect.top}px`;
                clone.style.width = '60px'; 
                clone.style.height = '90px';
                clone.style.borderRadius = '8px';
                clone.style.zIndex = '9999';
                clone.style.transition = 'all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275)';
                clone.style.pointerEvents = 'none';
                document.body.appendChild(clone);
                
                setTimeout(() => {
                    clone.style.left = `50%`;
                    clone.style.top = `90%`; // ลอยมาทางผู้เล่นด้านล่างจอ
                    clone.style.transform = 'translate(-50%, -50%) scale(1.2)';
                    clone.style.opacity = '0';
                }, 10);
                
                setTimeout(() => clone.remove(), 400);
            }
        }
    } catch(e) { console.warn('Animation error ignored', e); }
    // -----------------------------------------------------------

    if(isHost) processPlayerAction(myPeerId, action, index);
    else hostConnection.send({ type: 'PLAYER_ACTION', playerId: myPeerId, action, cardIndex: index });
};

window.showFloatingScore = (targetId, amt) => {
    let target = document.getElementById(`avatar-players-status-bar-${targetId}`);
    if (target) {
        let floatEl = document.createElement('div'); 
        floatEl.className = 'floating-sc';
        
        // --- [เพิ่มมิติการแสดงผล] Visual Trigger: แอนิเมชันเหรียญที่เด้งขึ้น ---
        floatEl.style.transition = 'all 1.5s ease-out';
        floatEl.style.position = 'absolute';
        floatEl.style.top = '0px';
        floatEl.style.left = '50%';
        floatEl.style.transform = 'translate(-50%, 0) scale(0.5)';
        floatEl.style.opacity = '1';
        floatEl.style.fontWeight = 'bold';
        floatEl.style.fontSize = '1.2rem';
        floatEl.style.textShadow = '0 0 5px rgba(0,0,0,0.8)';
        
        floatEl.style.color = amt > 0 ? '#32cd32' : '#ff4d4d'; 
        floatEl.innerText = amt > 0 ? `+${amt}` : amt; 
        
        floatEl.setAttribute('aria-hidden', 'true'); 
        target.style.position = 'relative'; // Ensure relative positioning
        target.appendChild(floatEl); 
        
        // Trigger reflow & Animate
        void floatEl.offsetWidth; 
        setTimeout(() => {
            floatEl.style.transform = 'translate(-50%, -40px) scale(1.2)';
            floatEl.style.opacity = '0';
        }, 10);

        setTimeout(() => floatEl.remove(), 1500);
    }
};

window.showResultModal = (title, detail, isEnd, winners, losers, isLanguageRefresh = false) => {
    lastResultRawData = { title, detail, isEnd, winners, losers };

    if (!isLanguageRefresh) {
        const srOnlyAnnouncer = document.getElementById('sr-only-announcer'); if (srOnlyAnnouncer) srOnlyAnnouncer.innerText = '';
        const visAnnouncer = document.getElementById('visible-game-announcer'); if (visAnnouncer) visAnnouncer.innerText = '';
        window.isStartingRound = false;
    }

    let resolvedTitle = localizeDynamicMessage(title);
    let resolvedDetail = detail
        .split('<br>')
        .map(line => localizeDynamicMessage(line))
        .join('<br>');

    document.getElementById('result-title').innerText = resolvedTitle;
    document.getElementById('result-details').innerHTML = resolvedDetail;
    
    document.getElementById('result-modal').style.display = 'flex';
    updateHomeBtnVisibility();

    if (winners && losers) {
        winners.forEach(id => { let el = document.getElementById(`avatar-result-status-bar-${id}`); if (el) el.classList.add('anim-happy'); });
        losers.forEach(id => { let el = document.getElementById(`avatar-result-status-bar-${id}`); if (el) el.classList.add('anim-sad'); });
    }

    let btnNext = document.getElementById('btn-next-round'); let waitTxt = document.getElementById('guest-waiting-next-round');
    
    let hasZeroCoins = (gameState.players && gameState.players.some(p => p.points === 0)) ||
                       (window.currentPublicState && window.currentPublicState.playersInfo && window.currentPublicState.playersInfo.some(p => p.points === 0));

    if (btnNext) {
        if (isHost && hasZeroCoins) {
            btnNext.style.display = 'none';
            if (waitTxt) waitTxt.style.display = 'none';
        } else if (isEnd) {
            btnNext.style.display = 'block'; btnNext.disabled = false; btnNext.innerText = t('btnBackHomeReload'); btnNext.onclick = () => window.location.reload();
            if (waitTxt) waitTxt.style.display = 'none';
        } else {
            if (isHost) {
                btnNext.style.display = 'block'; btnNext.disabled = false; btnNext.innerText = t('btnNextRound'); btnNext.onclick = closeResultModal;
                if (waitTxt) waitTxt.style.display = 'none';
            } else {
                btnNext.style.display = 'none';
                if (waitTxt) waitTxt.style.display = 'block';
            }
        }
    }

    if (!isLanguageRefresh) {
        if(myPeerId) {
            if (winners && winners.includes(myPeerId)) playSound('win');
            else if (losers && losers.includes(myPeerId)) playSound('lost');
            else playSound('no'); 
        }

        setTimeout(() => {
            if (!isLocalTempBotModalOpen) {
                const resultTitle = document.getElementById('result-title');
                if (resultTitle) resultTitle.focus();
            }
        }, 200);
    }
};

// =========================================================
// --- CHEAT MODE LOGIC (รวมฟังก์ชันไว้ล่างสุด) ---
// =========================================================
function setCheatHand(type) {
    let newHand = [];
    if (type === "ตอง") {
        newHand = [
            { suit: '♠', rank: '7', value: 5, val: 5, id: '7♠' },
            { suit: '♥', rank: '7', value: 5, val: 5, id: '7♥' },
            { suit: '♦', rank: '7', value: 5, val: 5, id: '7♦' },
            { suit: '♣', rank: '7', value: 5, val: 5, id: '7♣' },
            { suit: '♠', rank: 'K', value: 10, val: 10, id: 'K♠' },
            { suit: '♥', rank: 'K', value: 10, val: 10, id: 'K♥' },
            { suit: '♦', rank: 'K', value: 10, val: 10, id: 'K♦' }
        ];
    } else if (type === "ดอก" || type === "สี") {
        newHand = [
            { suit: '♠', rank: '2', value: 5, val: 5, id: '2♠' },
            { suit: '♠', rank: '3', value: 5, val: 5, id: '3♠' },
            { suit: '♠', rank: '4', value: 5, val: 5, id: '4♠' },
            { suit: '♠', rank: '5', value: 5, val: 5, id: '5♠' },
            { suit: '♠', rank: '6', value: 5, val: 5, id: '6♠' },
            { suit: '♠', rank: '7', value: 5, val: 5, id: '7♠' },
            { suit: '♠', rank: '8', value: 5, val: 5, id: '8♠' }
        ];
    } else if (type === "เรียง") {
        newHand = [
            { suit: '♥', rank: '4', value: 5, val: 5, id: '4♥' },
            { suit: '♥', rank: '5', value: 5, val: 5, id: '5♥' },
            { suit: '♥', rank: '6', value: 5, val: 5, id: '6♥' },
            { suit: '♣', rank: '8', value: 5, val: 5, id: '8♣' },
            { suit: '♣', rank: '9', value: 5, val: 5, id: '9♣' },
            { suit: '♣', rank: '10', value: 10, val: 10, id: '10♣' },
            { suit: '♣', rank: 'J', value: 10, val: 10, id: 'J♣' }
        ];
    } else if (type === "50") {
        newHand = [
            { suit: '♠', rank: '2', value: 50, val: 50, id: '2♠' },
            { suit: '♣', rank: 'Q', value: 50, val: 50, id: 'Q♣' },
            { suit: '♠', rank: 'A', value: 15, val: 15, id: 'A♠' },
            { suit: '♥', rank: 'A', value: 15, val: 15, id: 'A♥' },
            { suit: '♦', rank: 'A', value: 15, val: 15, id: 'A♦' },
            { suit: '♦', rank: '10', value: 10, val: 10, id: '10♦' },
            { suit: '♥', rank: 'K', value: 10, val: 10, id: 'K♥' }
        ];
    } else if (type === "3a") {
        newHand = [
            { suit: '♠', rank: 'A', value: 15, val: 15, id: 'A♠' },
            { suit: '♥', rank: 'A', value: 15, val: 15, id: 'A♥' },
            { suit: '♦', rank: 'A', value: 15, val: 15, id: 'A♦' },
            { suit: '♠', rank: '2', value: 50, val: 50, id: '2♠' },
            { suit: '♣', rank: 'Q', value: 50, val: 50, id: 'Q♣' },
            { suit: '♦', rank: 'K', value: 10, val: 10, id: 'K♦' },
            { suit: '♣', rank: 'K', value: 10, val: 10, id: 'K♣' }
        ];
    }

    if (isHost) {
        let p = gameState.players.find(x => x.id === myPeerId);
        if (p) {
            p.hand = newHand;
            syncStateToAll();
            announce(currentLang === 'en' ? 'Cheat hand applied (Host)' : 'เสกไพ่สำเร็จ (Host)', false);
        }
    } else {
        localPlayerState.hand = newHand;
        if (window.currentPublicState) {
            renderClientGame(window.currentPublicState, localPlayerState);
        }
        announce(currentLang === 'en' ? 'Cheat hand applied (Client)' : 'เสกไพ่สำเร็จ (Client)', false);
    }
}