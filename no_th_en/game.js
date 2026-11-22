import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { getDatabase, ref, set, get, onValue, update, remove, child, onDisconnect } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-database.js";

const firebaseConfig = {
    apiKey: "AIzaSyBIPeN4YUIrwM4MmFtDvT2DJ-v4toxc7tY",
    authDomain: "xo-game-134ec.firebaseapp.com",
    databaseURL: "https://xo-game-134ec-default-rtdb.asia-southeast1.firebasedatabase.app",
    projectId: "xo-game-134ec",
    storageBucket: "xo-game-134ec.firebasestorage.app",
    messagingSenderId: "318375224157",
    appId: "1:318375224157:web:1ae37e5d1f3c6bc2b77eb4"
};

const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

// ==================== Bilingual Support (TH / EN) ====================
let currentLang = localStorage.getItem('nothanks_lang') === 'en' ? 'en' : 'th';

const NOTHANKS_I18N = {
    th: {
        back_community_aria: "กลับสู่ชุมชน",
        main_subtitle: "การ์ดเกมนี้สำหรับการสันทนาการร่วมกัน",
        main_a11y_desc: "เกมนี้ออกแบบให้ผู้ใช้ Screen Reader และผู้เล่นทั่วไป สนุกร่วมกันได้อย่างเท่าเทียม",
        btn_create_room: "สร้างห้อง",
        btn_create_room_aria: "สร้างห้องใหม่",
        btn_how_to_play: "วิธีเล่น",
        btn_how_to_play_aria: "วิธีเล่น",
        open_rooms_heading: "รายชื่อห้องที่เปิดอยู่",
        room_list_region_aria: "รายการห้อง",
        searching_rooms: "กำลังค้นหาห้อง...",
        rules_title: "วิธีเล่น",
        rules_html: `
            <h3 tabindex="-1">เป้าหมายของเกม</h3>
            <p>ผู้เล่นต้องพยายามทำคะแนนให้น้อยที่สุดในตอนจบเกม</p>
            
            <h3 tabindex="-1">การเตรียมเกม</h3>
            <p>เกมนี้ใช้ไพ่หมายเลข 3 ถึง 35 (สุ่มนำไพ่ออก 9 ใบตอนเริ่มเกม)<br>ผู้เล่นแต่ละคนจะได้รับเหรียญเริ่มต้น (ขึ้นอยู่กับจำนวนผู้เล่น)</p>
            
            <h3 tabindex="-1">วิธีเล่นในแต่ละตา</h3>
            <p>เมื่อถึงตาของคุณ คุณมี 2 ทางเลือก:</p>
            <ul>
                <li><strong>ผ่าน (No Thanks!):</strong> จ่าย 1 เหรียญวางไว้บนไพ่ตรงกลาง และผ่านตานี้ไปให้คนถัดไป</li>
                <li><strong>รับไพ่:</strong> รับไพ่ใบนั้นมาไว้ในมือ และรับเหรียญทั้งหมดที่วางอยู่บนไพ่มาเป็นของคุณ จากนั้นจะมีการเปิดไพ่ใบใหม่ และคุณจะได้เล่นต่ออีกครั้ง</li>
            </ul>
            <p><em>หมายเหตุ: หากเหรียญหมด คุณไม่สามารถเลือก "ผ่าน" ได้ ต้อง "รับไพ่" เท่านั้น</em></p>
            
            <h3 tabindex="-1">การคิดคะแนน</h3>
            <ul>
                <li>ไพ่ 1 ใบ มีคะแนนเท่ากับตัวเลขบนไพ่ (ไพ่เยอะ = คะแนนเยอะ)</li>
                <li><strong>พิเศษ:</strong> หากคุณมีไพ่ที่ตัวเลขเรียงติดกัน (เช่น 12, 13, 14) จะคิดคะแนนเฉพาะ <strong>ไพ่ที่ใบต่ำสุด</strong> เท่านั้น (ในตัวอย่างคือ 12 คะแนน)</li>
                <li>เหรียญที่เหลือในตอนจบเกม 1 เหรียญ จะถูกนำไป <strong>ลบออกจากคะแนนรวม</strong> 1 คะแนน</li>
            </ul>
            <p><strong>ตัวอย่าง:</strong> คุณมีไพ่ 5, 8, 9, 10 และมีเหรียญ 3 เหรียญ<br>
            คะแนนไพ่ของคุณคือ 5 + 8 = 13 คะแนน<br>
            หักลบด้วยเหรียญ 3 เหรียญ = คุณได้คะแนนสุทธิ <strong>10 คะแนน</strong></p>
        `,
        btn_back_main: "กลับเมนูหลัก",
        lobby_title: "ห้องรอเริ่มเกม",
        room_id_label: "รหัสห้อง:",
        players_count_label: "ผู้เล่น:",
        choose_color_heading: "เลือกสีตัวละครของคุณ",
        choose_color_aria: "เลือกสีตัวละคร",
        player_list_heading: "รายชื่อผู้เล่น",
        btn_add_bot_prefix: "+ บอท",
        btn_add_bot_aria: "เพิ่มบอท 1 ตัว",
        btn_remove_bot: "- ลบบอท",
        btn_remove_bot_aria: "ลบบอทออก 1 ตัว",
        btn_start_game: "เริ่มเกม",
        btn_start_game_aria: "เริ่มเกม",
        client_wait_msg: "รอหัวหน้าห้องเริ่มเกม...",
        btn_leave_room: "ออกจากห้อง",
        btn_leave_room_aria: "ออกจากห้อง",
        playing_game_title: "กำลังเล่นเกม",
        btn_refresh_audio_aria: "รีเฟรชระบบเสียง",
        game_board_aria: "กระดานกลาง",
        table_coins_label: "เหรียญบนไพ่:",
        game_controls_aria: "ตัวเลือกการเล่นของคุณ",
        btn_pass: "ผ่าน (-1 🪙)",
        btn_pass_aria: "ผ่าน จ่าย 1 เหรียญ",
        btn_take: "รับไพ่",
        btn_take_aria: "รับไพ่ และรับเหรียญบนโต๊ะ",
        all_players_title: "สถานะผู้เล่นทุกคน",
        btn_close: "ปิด",
        btn_view_others: "👥 สถานะผู้เล่นทุกคน",
        btn_view_others_aria: "ดูสถานะผู้เล่นทุกคน",
        result_title: "สรุปผลคะแนน"
    },
    en: {
        back_community_aria: "Back to Community",
        main_subtitle: "Card game for shared recreation and fun",
        main_a11y_desc: "Designed for Screen Reader users and sighted players to enjoy equally together.",
        btn_create_room: "Create Room",
        btn_create_room_aria: "Create a new room",
        btn_how_to_play: "How to Play",
        btn_how_to_play_aria: "How to play",
        open_rooms_heading: "Open Rooms",
        room_list_region_aria: "Room list",
        searching_rooms: "Searching for rooms...",
        rules_title: "How to Play",
        rules_html: `
            <h3 tabindex="-1">Objective</h3>
            <p>Try to finish the game with the lowest score possible.</p>
            
            <h3 tabindex="-1">Setup</h3>
            <p>The deck has cards numbered 3 to 35 (9 random cards are removed at the start).<br>Each player starts with a set number of coins depending on player count.</p>
            
            <h3 tabindex="-1">On Your Turn</h3>
            <p>When it is your turn, you have 2 choices:</p>
            <ul>
                <li><strong>Pass (No Thanks!):</strong> Place 1 coin onto the center card and pass the turn to the next player.</li>
                <li><strong>Take Card:</strong> Take the center card into your hand along with all coins on it. A new card is then revealed and you take another turn.</li>
            </ul>
            <p><em>Note: If you have 0 coins left, you cannot pass and must take the card.</em></p>
            
            <h3 tabindex="-1">Scoring</h3>
            <ul>
                <li>Each card adds its face value to your score (higher = worse).</li>
                <li><strong>Special:</strong> For consecutive runs of cards (e.g. 12, 13, 14), only the <strong>lowest card</strong> counts (12 points in this example).</li>
                <li>Each remaining coin at the end of the game <strong>subtracts 1 point</strong> from your total score.</li>
            </ul>
            <p><strong>Example:</strong> You have cards 5, 8, 9, 10 and 3 coins.<br>
            Your card points are 5 + 8 = 13.<br>
            Minus 3 coins = net score of <strong>10 points</strong>.</p>
        `,
        btn_back_main: "Back to Main Menu",
        lobby_title: "Waiting Lobby",
        room_id_label: "Room ID:",
        players_count_label: "Players:",
        choose_color_heading: "Choose Your Character Color",
        choose_color_aria: "Choose character color",
        player_list_heading: "Players",
        btn_add_bot_prefix: "+ Bot",
        btn_add_bot_aria: "Add 1 bot",
        btn_remove_bot: "- Remove Bot",
        btn_remove_bot_aria: "Remove 1 bot",
        btn_start_game: "Start Game",
        btn_start_game_aria: "Start game",
        client_wait_msg: "Waiting for host to start the game...",
        btn_leave_room: "Leave Room",
        btn_leave_room_aria: "Leave room",
        playing_game_title: "Playing Game",
        btn_refresh_audio_aria: "Refresh audio system",
        game_board_aria: "Center board",
        table_coins_label: "Coins on card:",
        game_controls_aria: "Your game actions",
        btn_pass: "Pass (-1 🪙)",
        btn_pass_aria: "Pass and pay 1 coin",
        btn_take: "Take Card",
        btn_take_aria: "Take card and all coins on table",
        all_players_title: "All Players Status",
        btn_close: "Close",
        btn_view_others: "👥 All Players Status",
        btn_view_others_aria: "View all players status",
        result_title: "Score Summary"
    }
};

const COLOR_NAME_EN = {
    red: "Red",
    blue: "Blue",
    green: "Green",
    yellow: "Yellow",
    purple: "Purple",
    orange: "Orange",
    pink: "Pink"
};

function getColorName(color) {
    if (!color) return "";
    if (currentLang === 'en') {
        return COLOR_NAME_EN[color.id] || color.name;
    }
    return color.name;
}

function translateAnnouncementToEn(msg) {
    if (!msg) return '';
    let s = String(msg);

    if (s === 'ห้องเต็มหรือเกมเริ่มไปแล้ว') return 'Room is full or the game has already started.';
    if (s === 'ไพ่หมดกองแล้ว จบเกม! กำลังสรุปผลคะแนน') return 'The deck is empty. Game over! Calculating final scores.';

    s = s.replace(
        /^เริ่มเกมแล้ว! เปิดไพ่ใบแรก เลข (\d+) เป็นรอบของ (.+)$/,
        'Game started! First card is $1. It is $2\'s turn.'
    );
    s = s.replace(
        /^(.+?) เลือกผ่าน จ่าย 1 เหรียญ\. เหรียญบนไพ่รวมเป็น (\d+)$/,
        '$1 passed and paid 1 coin. Total coins on card: $2.'
    );
    s = s.replace(/^รอบของ(.+)$/, '$1\'s turn.');
    s = s.replace(
        /^(.+?) รับไพ่เลข (\d+) และรับเหรียญ (\d+) เหรียญ$/,
        '$1 took card $2 and gained $3 coins.'
    );
    s = s.replace(
        /^เปิดไพ่ใบใหม่ เลข (\d+)\. รอบของ(.+?) อีกครั้ง$/,
        'New card revealed: $1. $2\'s turn again.'
    );

    return s;
}

function applyStaticLanguage() {
    const dict = NOTHANKS_I18N[currentLang] || NOTHANKS_I18N.th;
    document.documentElement.lang = currentLang;

    document.querySelectorAll('[data-i18n]').forEach(el => {
        const k = el.getAttribute('data-i18n');
        if (dict[k] !== undefined) el.textContent = dict[k];
    });

    document.querySelectorAll('[data-i18n-html]').forEach(el => {
        const k = el.getAttribute('data-i18n-html');
        if (dict[k] !== undefined) el.innerHTML = dict[k];
    });

    document.querySelectorAll('[data-i18n-aria]').forEach(el => {
        const k = el.getAttribute('data-i18n-aria');
        if (dict[k] !== undefined) el.setAttribute('aria-label', dict[k]);
    });

    const label = currentLang === 'th' ? '🌐 EN' : '🌐 ไทย';
    const aria = currentLang === 'th' ? 'Switch language to English' : 'สลับภาษาเป็นภาษาไทย';
    ['btn-lang-main', 'btn-lang-lobby'].forEach(id => {
        const b = document.getElementById(id);
        if (b) {
            b.textContent = label;
            b.setAttribute('aria-label', aria);
        }
    });
}

function toggleLanguage() {
    currentLang = currentLang === 'th' ? 'en' : 'th';
    localStorage.setItem('nothanks_lang', currentLang);
    applyStaticLanguage();

    if (document.getElementById('screen-main').classList.contains('active')) {
        initRoomListener();
    }
    if (document.getElementById('screen-lobby').classList.contains('active')) {
        updateLobbyUI();
    }
    if (document.getElementById('screen-game').classList.contains('active')) {
        updateGameUI();
    }
    if (document.getElementById('screen-result').classList.contains('active')) {
        updateResultUI();
    }
    announce(currentLang === 'en' ? 'Language switched to English' : 'เปลี่ยนภาษาเป็นภาษาไทยแล้ว');
}

// --- ระบบจัดการ Web Audio API ---
let audioCtx = null;
const audioBuffers = {};
let bgmSource = null;

const audioFiles = {
    '1': 'audio/1.mp3',
    'select': 'audio/select.mp3',
    'start': 'audio/start.mp3',
    'bgm': 'audio/bgm.mp3',
    'jua': 'audio/jua.mp3',
    'skip': 'audio/skip.mp3',
    'win': 'audio/win.mp3',
    'no': 'audio/no.mp3',
    'lost': 'audio/lost.mp3'
};

function getAudioContext() {
    if (!audioCtx) {
        audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    }
    if (audioCtx.state === 'suspended') {
        audioCtx.resume();
    }
    return audioCtx;
}

async function loadAudio(key, url) {
    try {
        const response = await fetch(url);
        const arrayBuffer = await response.arrayBuffer();
        const ctx = getAudioContext();
        const audioBuffer = await ctx.decodeAudioData(arrayBuffer);
        audioBuffers[key] = audioBuffer;
    } catch (e) {
        console.warn(`Failed to load audio ${key}:`, e);
    }
}

function initAudio() {
    for (const [key, url] of Object.entries(audioFiles)) {
        loadAudio(key, url);
    }
}

function playSound(key, loop = false) {
    try {
        const ctx = getAudioContext();
        if (!audioBuffers[key]) return null;
        
        const source = ctx.createBufferSource();
        source.buffer = audioBuffers[key];
        source.loop = loop;
        source.connect(ctx.destination);
        source.start(0);
        return source;
    } catch (e) {
        console.warn("Play sound error:", e);
        return null;
    }
}

function stopBGM() {
    if (bgmSource) {
        try {
            bgmSource.stop();
        } catch(e){}
        bgmSource = null;
    }
}

function playBGM() {
    stopBGM();
    bgmSource = playSound('bgm', true);
}

function broadcastSound(soundKey) {
    playSound(soundKey);
    if (isHost) {
        connections.forEach(conn => {
            if (conn.open) conn.send({ type: 'playSound', sound: soundKey });
        });
    }
}

window.addEventListener('click', () => { getAudioContext(); }, { once: false });
initAudio();

const PLAYER_COLORS = [
    { id: "red", name: "สีแดง", hex: "#f44336" },
    { id: "blue", name: "สีน้ำเงิน", hex: "#2196F3" },
    { id: "green", name: "สีเขียว", hex: "#4CAF50" },
    { id: "yellow", name: "สีเหลือง", hex: "#FFEB3B" },
    { id: "purple", name: "สีม่วง", hex: "#9C27B0" },
    { id: "orange", name: "สีส้ม", hex: "#FF9800" },
    { id: "pink", name: "สีชมพู", hex: "#E91E63" }
];

let myPeerId = null;
let myColor = null;
let isHost = false;
let currentRoomId = null;
let heartbeatInterval = null;

let peer = null;
let connections = []; 
let hostConnection = null; 

let players = []; 
let game = {
    status: 'waiting', 
    deck: [],
    tableCard: null,
    tableCoins: 0,
    turnIndex: 0,
    playerStates: {},
    turnStartTime: null
};

let isJoiningRoom = false;
let isCreatingRoom = false; 
let joinTimeoutId = null;
let myLastTurnState = null;
let turnAnnounceTimeout = null;

let turnTimer = null;

// --- ตัวแปรสำหรับ Visual Trigger ---
let visualStateMap = {
    coins: -1,
    cards: null
};

// --- Visual Background Init ---
function initBackground() {
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    const container = document.getElementById('bg-decor');
    if(!container) return;
    container.innerHTML = '';
    for(let i=0; i<15; i++) {
        let b = document.createElement('div');
        b.className = 'bg-bubble';
        b.style.width = Math.random() * 80 + 20 + 'px';
        b.style.height = b.style.width;
        b.style.left = Math.random() * 100 + '%';
        b.style.animationDuration = Math.random() * 15 + 10 + 's';
        b.style.animationDelay = Math.random() * 5 + 's';
        container.appendChild(b);
    }
}
window.addEventListener('DOMContentLoaded', () => {
    initBackground();
    applyStaticLanguage();
    const btnLangMain = document.getElementById('btn-lang-main');
    const btnLangLobby = document.getElementById('btn-lang-lobby');
    if (btnLangMain) btnLangMain.addEventListener('click', toggleLanguage);
    if (btnLangLobby) btnLangLobby.addEventListener('click', toggleLanguage);
});

function stopTimer() {
    if (turnTimer) {
        clearInterval(turnTimer);
        turnTimer = null;
    }
    const container = document.getElementById('timer-bar-container');
    if (container) {
        container.style.display = 'none';
    }
}

function startOrUpdateTimer() {
    stopTimer();

    if (game.status !== 'playing') return;
    const currentPlayer = players[game.turnIndex];
    if (!currentPlayer || currentPlayer.isBot) return;

    const container = document.getElementById('timer-bar-container');
    if (!container) return;

    container.style.display = 'block';

    function tick() {
        if (game.status !== 'playing') {
            stopTimer();
            return;
        }
        const now = Date.now();
        const elapsed = Math.floor((now - (game.turnStartTime || now)) / 1000);
        const remaining = Math.max(0, 40 - elapsed);

        const textEl = document.getElementById('timer-bar-text');
        const fillEl = document.getElementById('timer-bar-fill');

        if (textEl) {
            textEl.textContent = currentLang === 'en'
                ? `${remaining}s left`
                : `เหลือเวลา ${remaining} วินาที`;
        }

        if (fillEl) {
            const percentage = Math.max(0, Math.min(100, (remaining / 40) * 100));
            fillEl.style.width = `${percentage}%`;

            if (remaining > 20) {
                fillEl.style.backgroundColor = 'var(--primary)';
            } else if (remaining > 10) {
                fillEl.style.backgroundColor = '#ff9800';
            } else {
                fillEl.style.backgroundColor = 'var(--danger)';
            }
        }

        if (remaining <= 0) {
            stopTimer();
            if (isHost && game.status === 'playing') {
                const pState = game.playerStates[currentPlayer.id];
                if (pState) {
                    if (pState.coins > 0) {
                        processAction(currentPlayer, 'pass');
                    } else {
                        processAction(currentPlayer, 'take');
                    }
                }
            }
        }
    }

    tick();
    turnTimer = setInterval(tick, 200);
}

function startHeartbeat() {
    stopHeartbeat();
    heartbeatInterval = setInterval(() => {
        if (isHost && currentRoomId && game.status === 'waiting') {
            update(ref(db, `nothanks_rooms/${currentRoomId}`), {
                lastActive: Date.now()
            });
        }
    }, 5000);
}

function stopHeartbeat() {
    if (heartbeatInterval) {
        clearInterval(heartbeatInterval);
        heartbeatInterval = null;
    }
}

function updateRoomInFirebase() {
    if (!isHost || !currentRoomId || game.status !== 'waiting') return;
    const roomRef = ref(db, `nothanks_rooms/${currentRoomId}`);
    update(roomRef, {
        currentPlayers: players.length,
        maxPlayers: 7,
        playerNames: players.map(p => getPronoun(p)),
        lastActive: Date.now()
    });
}

// --- Visual Animation Hooks ---
function visualAnimThrowCoin(playerId) {
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    let startEl = document.getElementById(`vis-avatar-${playerId}`);
    if (playerId === myPeerId) {
        startEl = document.querySelector('#my-status-group .hud-avatar');
    }
    const tableCoin = document.getElementById('table-coins-display');
    const tableCard = document.getElementById('table-card-display');
    if (!startEl || !tableCoin) return;

    const startRect = startEl.getBoundingClientRect();
    const endRect = tableCoin.getBoundingClientRect();

    const coin = document.createElement('div');
    coin.className = 'visual-fly-coin';
    coin.textContent = '🪙';
    coin.setAttribute('aria-hidden', 'true');
    document.body.appendChild(coin);

    coin.style.left = `${startRect.left + startRect.width/2 - 16}px`;
    coin.style.top = `${startRect.top + startRect.height/2 - 16}px`;

    requestAnimationFrame(() => {
        coin.style.transform = `translate(${endRect.left - startRect.left}px, ${endRect.top - startRect.top}px) scale(0.5)`;
        coin.style.opacity = '0';
    });

    setTimeout(() => {
        if (coin.parentNode) coin.parentNode.removeChild(coin);
        tableCoin.parentElement.classList.add('pulse-anim');
        if (tableCard) tableCard.classList.add('coin-hit-anim');
        setTimeout(() => {
            tableCoin.parentElement.classList.remove('pulse-anim');
            if (tableCard) tableCard.classList.remove('coin-hit-anim');
        }, 400);
    }, 500);
}

function visualAnimDealCard(cardNum) {
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    const deck = document.getElementById('deck-info-container');
    const tableCard = document.getElementById('table-card-display');
    if (!deck || !tableCard) return;

    const startRect = deck.getBoundingClientRect();
    const endRect = tableCard.getBoundingClientRect();

    const card = document.createElement('div');
    card.className = 'visual-fly-card';
    card.textContent = cardNum;
    card.setAttribute('aria-hidden', 'true');
    document.body.appendChild(card);

    card.style.left = `${startRect.left + startRect.width/2 - 40}px`;
    card.style.top = `${startRect.top + startRect.height/2 - 55}px`;
    card.style.transform = 'rotateY(180deg) scale(0.5)';
    card.style.transition = 'transform 0.7s cubic-bezier(0.34, 1.56, 0.64, 1), opacity 0.7s';

    requestAnimationFrame(() => {
        card.style.transform = `translate(${endRect.left - startRect.left}px, ${endRect.top - startRect.top}px) rotateY(0deg) scale(1)`;
        card.style.opacity = '1';
    });

    tableCard.style.opacity = '0';

    setTimeout(() => {
        if (card.parentNode) card.parentNode.removeChild(card);
        tableCard.style.opacity = '1';
        tableCard.classList.add('pop-anim');
        setTimeout(() => tableCard.classList.remove('pop-anim'), 400);
    }, 700);
}

function visualAnimTakeCard(playerId, cardNum, coins) {
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    let endEl = document.getElementById(`vis-avatar-${playerId}`);
    if (playerId === myPeerId) {
        endEl = document.querySelector('#my-status-group .hud-avatar');
    }
    const tableCard = document.getElementById('table-card-display');
    if (!endEl || !tableCard) return;

    const endRect = endEl.getBoundingClientRect();
    const startRect = tableCard.getBoundingClientRect();

    // Card Visual Move
    const card = document.createElement('div');
    card.className = 'visual-fly-card';
    card.textContent = cardNum;
    card.setAttribute('aria-hidden', 'true');
    document.body.appendChild(card);

    card.style.left = `${startRect.left + startRect.width/2 - 40}px`;
    card.style.top = `${startRect.top + startRect.height/2 - 55}px`;
    card.style.transition = 'transform 0.6s cubic-bezier(0.34, 1.56, 0.64, 1), opacity 0.6s';

    requestAnimationFrame(() => {
        card.style.transform = `translate(${endRect.left - startRect.left}px, ${endRect.top - startRect.top}px) scale(0.3) rotate(15deg)`;
        card.style.opacity = '0.2';
    });

    tableCard.style.opacity = '0';

    // Coins Visual Fly
    for (let i = 0; i < coins; i++) {
        const coin = document.createElement('div');
        coin.className = 'visual-fly-coin';
        coin.textContent = '🪙';
        coin.setAttribute('aria-hidden', 'true');
        document.body.appendChild(coin);
        
        coin.style.left = `${startRect.left + startRect.width/2 - 16}px`;
        coin.style.top = `${startRect.top + startRect.height/2 - 16}px`;

        const scatterX = (Math.random() - 0.5) * 80;
        const scatterY = (Math.random() - 0.5) * 80;

        requestAnimationFrame(() => {
            coin.style.transition = `transform 0.5s cubic-bezier(0.25, 1, 0.5, 1) ${i * 0.02}s, opacity 0.5s ${i * 0.02}s`;
            coin.style.transform = `translate(${endRect.left - startRect.left + scatterX}px, ${endRect.top - startRect.top + scatterY}px) scale(0.5)`;
            coin.style.opacity = '0';
        });

        setTimeout(() => { if (coin.parentNode) coin.parentNode.removeChild(coin); }, 600 + i * 20);
    }

    setTimeout(() => {
        if (card.parentNode) card.parentNode.removeChild(card);
    }, 600);
}

function triggerVisualsFromAnnounce(rawMessage) {
    if (typeof rawMessage !== 'string') return;
    let passMatch = rawMessage.match(/\[PLAYER:(.+?)\] เลือกผ่าน/);
    if (passMatch) { visualAnimThrowCoin(passMatch[1]); return; }
    let takeMatch = rawMessage.match(/\[PLAYER:(.+?)\] รับไพ่เลข (\d+) และรับเหรียญ (\d+)/);
    if (takeMatch) { visualAnimTakeCard(takeMatch[1], takeMatch[2], takeMatch[3]); return; }
    let newCardMatch = rawMessage.match(/เปิดไพ่ใบใหม่ เลข (\d+)/);
    if (newCardMatch) { visualAnimDealCard(newCardMatch[1]); return; }
    let startMatch = rawMessage.match(/เริ่มเกมแล้ว! เปิดไพ่ใบแรก เลข (\d+)/);
    if (startMatch) { visualAnimDealCard(startMatch[1]); return; }
}

// --- ระบบ Accessibility Announcements ---
function announce(message, assertive = false) {
    let originalMessage = message; 
    if (typeof message === 'string') {
        message = message.replace(/\[PLAYER:(.+?)\]/g, (match, id) => {
            const p = players.find(p => p.id === id);
            return p ? getPronoun(p) : (currentLang === 'en' ? "Player" : "ผู้เล่น");
        });
        if (currentLang === 'en') {
            message = translateAnnouncementToEn(message);
        }
    }

    const targetId = assertive ? 'aria-assertive' : 'aria-polite';
    const el = document.getElementById(targetId);
    el.textContent = ''; 
    setTimeout(() => { 
        el.textContent = message; 
        setTimeout(() => {
            el.textContent = ''; 
        }, 3000);
    }, 50);
    
    // Hook to trigger visuals exactly synchronized with original game logic
    triggerVisualsFromAnnounce(originalMessage);
}

function switchScreen(screenId, focusId) {
    if (screenId !== 'screen-game') {
        stopTimer();
    }
    document.querySelectorAll('.screen').forEach(s => {
        s.classList.remove('active');
        s.classList.remove('view-all-mode'); 
    });
    document.getElementById('all-players-section').style.display = 'none'; 
    document.getElementById(screenId).classList.add('active');
    if (focusId) {
        document.getElementById(focusId).focus();
    }
}

function getPronoun(player) {
    const colorName = getColorName(player.color);
    if (currentLang === 'en') {
        if (player.peerId === myPeerId) return `You (${colorName})`;
        if (player.isBot) return `Bot (${colorName})`;
        return `Friend (${colorName})`;
    }
    if (player.peerId === myPeerId) return `คุณ${colorName}`;
    if (player.isBot) return `บอท${colorName}`;
    return `เพื่อน${colorName}`;
}

function initRoomListener() {
    const roomsRef = ref(db, 'nothanks_rooms');
    const mainRoomList = document.getElementById('main-room-list');
    
    onValue(roomsRef, (snapshot) => {
        if (isJoiningRoom || currentRoomId) return; 
        const isEn = currentLang === 'en';
        
        let ul = mainRoomList.querySelector('ul.room-list');
        if (!ul) {
            mainRoomList.innerHTML = '';
            ul = document.createElement('ul');
            ul.className = 'room-list';
            mainRoomList.appendChild(ul);
        }

        if (snapshot.exists()) {
            const rooms = snapshot.val();
            let roomCount = 0;
            const now = Date.now();
            const currentRooms = [];
            const existingRooms = Array.from(ul.children).map(li => li.dataset.roomId);
            
            for (let roomId in rooms) {
                if (rooms[roomId].status === 'waiting') {
                    const lastActive = rooms[roomId].lastActive || rooms[roomId].timestamp || 0;
                    if (now - lastActive > 15000) {
                        remove(ref(db, `nothanks_rooms/${roomId}`));
                        continue;
                    }

                    roomCount++;
                    currentRooms.push(roomId);
                    
                    const currentP = rooms[roomId].currentPlayers || 1;
                    const maxP = rooms[roomId].maxPlayers || 7;
                    const btnText = isEn
                        ? `Join Room ${roomId} (${currentP}/${maxP})`
                        : `เข้าร่วมห้อง ${roomId} (${currentP}/${maxP})`;
                    const ariaText = isEn
                        ? `Join room ${roomId}, ${currentP} of ${maxP} players`
                        : `เข้าร่วมห้อง ${roomId} ผู้เล่น ${currentP} จาก ${maxP} คน`;
                    
                    let existingLi = ul.querySelector(`li[data-room-id="${roomId}"]`);
                    if (existingLi) {
                        let btn = existingLi.querySelector('button');
                        if (btn && !isJoiningRoom) {
                            btn.textContent = btnText;
                            btn.setAttribute("aria-label", ariaText);
                        }
                    } else {
                        let li = document.createElement('li');
                        li.dataset.roomId = roomId;
                        
                        let btn = document.createElement('button');
                        btn.textContent = btnText;
                        btn.setAttribute("aria-label", ariaText);
                        btn.style.width = "100%";
                        
                        if (isJoiningRoom) {
                            btn.disabled = true;
                        }
                        
                        btn.addEventListener('click', function() { 
                            joinRoom(roomId, rooms[roomId].hostId, this); 
                        });
                        
                        li.appendChild(btn);
                        ul.appendChild(li);
                    }
                }
            }

            existingRooms.forEach(id => {
                if (!currentRooms.includes(id)) {
                    let liToRemove = ul.querySelector(`li[data-room-id="${id}"]`);
                    if (liToRemove) ul.removeChild(liToRemove);
                }
            });

            if (roomCount === 0) {
                mainRoomList.innerHTML = `<p style="text-align:center;">${isEn ? 'No open rooms available right now' : 'ไม่มีห้องที่เปิดอยู่ในขณะนี้'}</p>`;
            }
        } else {
            mainRoomList.innerHTML = `<p style="text-align:center;">${isEn ? 'No open rooms available right now' : 'ไม่มีห้องที่เปิดอยู่ในขณะนี้'}</p>`;
        }
    });
}

async function getNextRoomId() {
    const counterRef = ref(db, 'nothanks_room_counter/last_id');
    const snapshot = await get(counterRef);
    let maxId = 0;
    if (snapshot.exists()) {
        maxId = snapshot.val();
    }
    const nextIdNum = maxId + 1;
    await set(counterRef, nextIdNum);
    
    const nextIdStr = nextIdNum.toString().padStart(5, '0');
    return `no${nextIdStr}`; 
}

function initPeer(onReady) {
    peer = new Peer();
    peer.on('open', (id) => {
        myPeerId = id;
        onReady();
    });
    peer.on('error', (err) => {
        announce(currentLang === 'en' ? "Network connection error" : "เกิดข้อผิดพลาดในการเชื่อมต่อเครือข่าย", true);
        if (isJoiningRoom) {
            resetJoinState();
        }
        if (isCreatingRoom) {
            isCreatingRoom = false;
            const btn = document.getElementById('btn-create-room');
            if (btn) btn.disabled = false;
            if (peer) { peer.destroy(); peer = null; }
        }
    });
}

document.getElementById('btn-create-room').addEventListener('click', () => {
    if (isCreatingRoom) return; 
    isCreatingRoom = true;
    
    const btnCreate = document.getElementById('btn-create-room');
    btnCreate.disabled = true; 

    playSound('1');
    announce(currentLang === 'en' ? "Creating room, please wait..." : "กำลังสร้างห้อง กรุณารอซักครู่");
    
    let isPeerCallbackCalled = false; 

    initPeer(async () => {
        if (isPeerCallbackCalled) return;
        isPeerCallbackCalled = true;

        try {
            isHost = true;
            myColor = PLAYER_COLORS[0]; 
            currentRoomId = await getNextRoomId();
            
            players = [{ id: myPeerId, peerId: myPeerId, color: myColor, isBot: false, online: true }];

            const roomRef = ref(db, `nothanks_rooms/${currentRoomId}`);
            await set(roomRef, {
                hostId: myPeerId,
                status: 'waiting',
                timestamp: Date.now(),
                lastActive: Date.now(),
                currentPlayers: players.length,
                maxPlayers: 7,
                playerNames: players.map(p => getPronoun(p))
            });
            onDisconnect(roomRef).remove();

            setupHostEvents();
            startHeartbeat();
            updateLobbyUI();
            switchScreen('screen-lobby', 'title-lobby');
            announce(currentLang === 'en'
                ? `Room created! Room ID is ${currentRoomId}. Your starting color is ${getColorName(myColor)}.`
                : `สร้างห้องสำเร็จ รหัสห้องคือ ${currentRoomId} คุณเริ่มต้นที่${myColor.name}`);
            
            isCreatingRoom = false; 
            btnCreate.disabled = false;
        } catch (error) {
            announce(currentLang === 'en' ? "Error creating room" : "เกิดข้อผิดพลาดในการสร้างห้อง", true);
            isCreatingRoom = false;
            btnCreate.disabled = false;
            if (peer) { peer.destroy(); peer = null; }
        }
    });
});

function setupHostEvents() {
    peer.on('connection', (conn) => {
        conn.on('open', () => {
            const usedColors = players.map(p => p.color.id);
            const availableColor = PLAYER_COLORS.find(c => !usedColors.includes(c.id));
            
            if (players.length >= 7 || game.status === 'playing' || !availableColor) {
                conn.send({ type: 'reject', message: 'ห้องเต็มหรือเกมเริ่มไปแล้ว' });
                setTimeout(() => conn.close(), 1000);
                return;
            }

            connections.push(conn);
            const newPlayer = { id: conn.peer, peerId: conn.peer, color: availableColor, isBot: false, online: true };
            players.push(newPlayer);
            
            announce(currentLang === 'en'
                ? `${getPronoun(newPlayer)} joined the room.`
                : `${getPronoun(newPlayer)} เข้าร่วมห้องแล้ว`);
            broadcastSound('select');
            broadcastState();

            conn.on('data', (data) => handleClientMessage(conn.peer, data));
            conn.on('close', () => handlePlayerDisconnect(conn.peer));
        });
    });
}

window.joinRoom = function(roomId, hostId, btn = null) {
    if (isJoiningRoom) return;
    isJoiningRoom = true;

    const allButtons = document.querySelectorAll('#main-room-list button');
    allButtons.forEach(b => {
        b.disabled = true;
        if (btn && b === btn) {
            b.textContent = currentLang === 'en' ? 'Joining room...' : 'กำลังเข้าร่วมห้อง...';
        }
    });

    announce(currentLang === 'en' ? `Connecting to room ${roomId}...` : `กำลังเชื่อมต่อไปยังห้อง ${roomId}`);
    
    joinTimeoutId = setTimeout(() => {
        if (isJoiningRoom) {
            announce(currentLang === 'en' ? "Connection failed. Please try again." : "เชื่อมต่อไม่สำเร็จ กรุณาลองใหม่อีกครั้ง", true);
            resetJoinState();
        }
    }, 10000);

    initPeer(() => {
        isHost = false;
        currentRoomId = roomId;
        hostConnection = peer.connect(hostId);
        
        hostConnection.on('open', () => {
            hostConnection.on('data', handleHostMessage);
            hostConnection.on('close', () => {
                if (isJoiningRoom) {
                    announce(currentLang === 'en'
                        ? "Host rejected the connection or the room was closed."
                        : "หัวหน้าห้องปฏิเสธการเชื่อมต่อ หรือห้องปิดไปแล้ว", true);
                    resetJoinState();
                } else {
                    announce(currentLang === 'en'
                        ? "Disconnected from host."
                        : "ขาดการเชื่อมต่อจากหัวหน้าห้อง", true);
                    leaveRoom();
                }
            });
        });
        hostConnection.on('error', (err) => {
            if (isJoiningRoom) {
                announce(currentLang === 'en' ? "Connection error" : "เกิดข้อผิดพลาดในการเชื่อมต่อ", true);
                resetJoinState();
            }
        });
    });
}

function resetJoinState() {
    isJoiningRoom = false;
    if (joinTimeoutId) {
        clearTimeout(joinTimeoutId);
        joinTimeoutId = null;
    }
    if (hostConnection) {
        hostConnection.close();
        hostConnection = null;
    }
    if (peer) {
        peer.destroy();
        peer = null;
    }
    currentRoomId = null;
    initRoomListener();
}

function broadcastState() {
    if (!isHost) return;
    const payload = { type: 'sync', players, game };
    connections.forEach(conn => {
        if(conn.open) conn.send(payload);
    });
    updateLobbyUI(); 
    if (game.status === 'waiting') {
        updateRoomInFirebase();
    } else if (game.status === 'playing') {
        updateGameUI();
    } else if (game.status === 'ended') {
        if (!document.getElementById('screen-result').classList.contains('active')) {
            switchScreen('screen-result', 'title-result');
        }
        updateResultUI();
    }
}

function handleHostMessage(data) {
    if (data.type === 'reject') {
        if (isJoiningRoom) resetJoinState();
        announce(data.message, true);
        if (!isJoiningRoom) leaveRoom();
    } else if (data.type === 'playSound') {
        playSound(data.sound);
    } else if (data.type === 'triggerStartAnim') {
        runStartAnim();
    } else if (data.type === 'triggerEndAnim') {
        players = data.players;
        game = data.game;
        runEndAnim();
    } else if (data.type === 'sync') {
        if (isJoiningRoom) {
            isJoiningRoom = false;
            if (joinTimeoutId) clearTimeout(joinTimeoutId);
        }

        const wasWaiting = game.status === 'waiting';
        const oldPlayers = [...players];
        
        players = data.players;
        game = data.game;
        
        const me = players.find(p => p.peerId === myPeerId);
        if(me) myColor = me.color;

        if (game.status === 'waiting') {
            const isEn = currentLang === 'en';
            if (!isHost && oldPlayers.length > 0) {
                const oldIds = oldPlayers.map(p => p.id);
                const newIds = players.map(p => p.id);
                
                players.forEach(p => {
                    if (!oldIds.includes(p.id) && p.peerId !== myPeerId) {
                        if (p.isBot) {
                            announce(isEn
                                ? `Bot (${getColorName(p.color)}) was added to the room.`
                                : `บอท${p.color.name} ถูกเพิ่มเข้าห้องแล้ว`);
                        } else {
                            announce(isEn
                                ? `Friend (${getColorName(p.color)}) joined the room.`
                                : `เพื่อน${p.color.name} เข้าร่วมห้องแล้ว`);
                        }
                    }
                });
                
                oldPlayers.forEach(p => {
                    if (!newIds.includes(p.id)) {
                        if (p.isBot) {
                            announce(isEn
                                ? `Bot (${getColorName(p.color)}) was removed from the room.`
                                : `บอท${p.color.name} ถูกลบออกจากห้องแล้ว`);
                        } else {
                            announce(isEn
                                ? `Friend (${getColorName(p.color)}) left the room.`
                                : `เพื่อน${p.color.name} ออกจากห้องแล้ว`);
                        }
                    }
                });
            }

            if (!document.getElementById('screen-lobby').classList.contains('active')) {
                switchScreen('screen-lobby', 'title-lobby');
                announce(isEn
                    ? `Joined room ${currentRoomId}! Your color is ${getColorName(myColor)}. Waiting for host to start.`
                    : `เข้าร่วมห้อง ${currentRoomId} สำเร็จ คุณได้รับ${myColor.name} รอหัวหน้าห้องเริ่มเกม`);
            }
            updateLobbyUI();
        } else if (game.status === 'playing') {
            if (wasWaiting) {
                switchScreen('screen-game', 'title-game');
            }
            updateGameUI();
        } else if (game.status === 'ended') {
            if (!document.getElementById('screen-result').classList.contains('active')) {
                switchScreen('screen-result', 'title-result');
            }
            updateResultUI();
        }
    } else if (data.type === 'announce') {
        announce(data.message, data.assertive);
    }
}

function handleClientMessage(peerId, data) {
    if (!isHost) return;
    if (data.type === 'action') {
        const playerIndex = players.findIndex(p => p.id === peerId);
        if (playerIndex === game.turnIndex) {
            processAction(players[playerIndex], data.action);
        }
    } else if (data.type === 'changeColor') {
        const isTaken = players.some(p => p.color.id === data.colorId);
        if (!isTaken) {
            const player = players.find(p => p.peerId === peerId);
            const newColor = PLAYER_COLORS.find(c => c.id === data.colorId);
            if (player && newColor) {
                player.color = newColor;
                broadcastState();
            }
        }
    }
}

function sendAction(action) {
    if (isHost) {
        processAction(players[game.turnIndex], action);
    } else {
        hostConnection.send({ type: 'action', action: action });
    }
}

function handlePlayerDisconnect(peerId) {
    if (!isHost) return;
    const p = players.find(p => p.peerId === peerId);
    if (p) {
        p.online = false;
        announce(currentLang === 'en'
            ? `${getPronoun(p)} disconnected. Replaced by a bot temporarily.`
            : `${getPronoun(p)} หลุดจากการเชื่อมต่อ เปลี่ยนเป็นบอทชั่วคราว`, true);
        p.isBot = true; 
        broadcastState();
        
        if (game.status === 'playing' && players[game.turnIndex].id === p.id) {
            setTimeout(() => processBotTurn(), 1500);
        }
    }
}

function updateLobbyUI() {
    document.getElementById('lobby-room-id').textContent = currentRoomId;
    document.getElementById('lobby-player-count').textContent = players.length;
    
    const ul = document.getElementById('lobby-players');
    ul.innerHTML = '';
    players.forEach(p => {
        let li = document.createElement('li');
        // Visual Enhancement: Avatars in list without breaking semantics
        li.innerHTML = `
            <div style="display:flex; align-items:center; gap:15px;">
                <div class="visual-avatar" aria-hidden="true" style="position:relative; transform:none; width:40px; height:40px; font-size:18px; background:${p.color.hex}"><span aria-hidden="true">^ᴗ^</span></div> 
                <span style="color:${p.color.hex}; font-size:18px; font-weight:bold;">${getPronoun(p)}</span>
            </div>
        `;
        li.setAttribute("role", "listitem");
        ul.appendChild(li);
    });

    renderColors();

    if (isHost) {
        document.getElementById('host-controls').style.display = 'block';
        document.getElementById('client-wait-msg').style.display = 'none';
        
        let botCount = players.filter(p => p.isBot).length;
        document.getElementById('bot-count-display').textContent = botCount;
        
        document.getElementById('btn-add-bot').disabled = players.length >= 7;
        document.getElementById('btn-remove-bot').disabled = botCount === 0;
        
        document.getElementById('btn-start-game').disabled = players.length < 3;
    } else {
        document.getElementById('host-controls').style.display = 'none';
        document.getElementById('client-wait-msg').style.display = 'block';
    }
}

function renderColors() {
    const container = document.getElementById('color-options');
    container.innerHTML = '';
    
    PLAYER_COLORS.forEach(c => {
        const label = document.createElement('label');
        label.style.cursor = 'pointer';
        
        const radio = document.createElement('input');
        radio.type = 'radio';
        radio.name = 'player-color';
        radio.value = c.id;
        radio.checked = (myColor && myColor.id === c.id);
        radio.className = 'sr-only'; // Kept accessible for SR and Keyboard
        
        const isTaken = players.some(p => p.id !== myPeerId && p.color.id === c.id);
        radio.disabled = isTaken;
        
        radio.addEventListener('change', () => {
            if (radio.checked) {
                if (isHost) {
                    myColor = c;
                    const me = players.find(p => p.peerId === myPeerId);
                    if (me) me.color = c;
                    broadcastState();
                } else {
                    hostConnection.send({ type: 'changeColor', colorId: c.id });
                }
            }
        });
        
        label.appendChild(radio);
        
        // Visual Layer
        const preview = document.createElement('div');
        preview.className = 'char-preview';
        preview.style.backgroundColor = isTaken ? '#555' : c.hex;
        preview.style.color = 'white';
        preview.innerHTML = '<span aria-hidden="true">^ᴗ^</span>';
        preview.setAttribute('aria-hidden', 'true');
        label.appendChild(preview);

        const textSpan = document.createElement('span');
        textSpan.className = 'char-name';
        textSpan.textContent = getColorName(c);
        textSpan.style.color = isTaken ? '#888' : '#fff';
        label.appendChild(textSpan);

        container.appendChild(label);
    });
}

document.getElementById('btn-add-bot').addEventListener('click', () => {
    if (players.length >= 7) return;
    const usedColors = players.map(p => p.color.id);
    const availableColor = PLAYER_COLORS.find(c => !usedColors.includes(c.id));
    
    const botId = `bot_${Date.now()}`;
    const newBot = { id: botId, peerId: null, color: availableColor, isBot: true, online: true };
    players.push(newBot);
    announce(currentLang === 'en'
        ? `Added ${getPronoun(newBot)}.`
        : `เพิ่ม${getPronoun(newBot)} เรียบร้อยแล้ว`);
    broadcastSound('select');
    broadcastState();
});

document.getElementById('btn-remove-bot').addEventListener('click', () => {
    for (let i = players.length - 1; i >= 0; i--) {
        if (players[i].isBot) {
            const removed = players.splice(i, 1)[0];
            announce(currentLang === 'en'
                ? `Removed ${getPronoun(removed)}.`
                : `ลบ${getPronoun(removed)} เรียบร้อยแล้ว`);
            break;
        }
    }
    broadcastSound('select');
    broadcastState();
});

function runStartAnim(onComplete) {
    stopBGM();
    playSound('start');
    
    const overlay = document.getElementById('anim-overlay');
    const textEl = document.getElementById('anim-text');
    overlay.style.display = 'flex';
    textEl.textContent = '';
    
    setTimeout(() => {
        textEl.textContent = 'No Thanks!';
        textEl.style.animation = 'none'; void textEl.offsetWidth; textEl.style.animation = 'bounceIn 0.8s cubic-bezier(0.175, 0.885, 0.32, 1.275)';
        announce('No Thanks!', true);
    }, 200);
    
    setTimeout(() => {
        textEl.textContent = '3';
        textEl.style.animation = 'none'; void textEl.offsetWidth; textEl.style.animation = 'bounceIn 0.8s cubic-bezier(0.175, 0.885, 0.32, 1.275)';
        announce('3', true);
    }, 1200);
    
    setTimeout(() => {
        textEl.textContent = '2';
        textEl.style.animation = 'none'; void textEl.offsetWidth; textEl.style.animation = 'bounceIn 0.8s cubic-bezier(0.175, 0.885, 0.32, 1.275)';
        announce('2', true);
    }, 2200);

    setTimeout(() => {
        textEl.textContent = '1';
        textEl.style.animation = 'none'; void textEl.offsetWidth; textEl.style.animation = 'bounceIn 0.8s cubic-bezier(0.175, 0.885, 0.32, 1.275)';
        announce('1', true);
    }, 3200);

    setTimeout(() => {
        textEl.textContent = currentLang === 'en' ? 'Start!' : 'เริ่มเกม!';
        textEl.style.animation = 'none'; void textEl.offsetWidth; textEl.style.animation = 'bounceIn 0.8s cubic-bezier(0.175, 0.885, 0.32, 1.275)';
        announce(currentLang === 'en' ? 'Take it or No Thanks!' : 'รับหรือไม่เอา', true);
    }, 4200);

    setTimeout(() => {
        overlay.style.display = 'none';
        textEl.textContent = '';
        document.getElementById('aria-polite').textContent = '';
        document.getElementById('aria-assertive').textContent = '';
        if (onComplete) onComplete();
        playBGM();
    }, 5090);
}

document.getElementById('btn-start-game').addEventListener('click', async () => {
    if (players.length < 3 || players.length > 7) return;
    
    const btnStart = document.getElementById('btn-start-game');
    btnStart.disabled = true;
    btnStart.style.display = 'none';

    connections.forEach(conn => {
        if (conn.open) conn.send({ type: 'triggerStartAnim' });
    });

    stopHeartbeat();
    await remove(ref(db, `nothanks_rooms/${currentRoomId}`));

    runStartAnim(() => {
        setupNewGame();
        switchScreen('screen-game', 'title-game');
        broadcastState();
        checkBotTurn();
    });
});

function setupNewGame() {
    let deck = Array.from({length: 33}, (_, i) => i + 3);
    deck = deck.sort(() => Math.random() - 0.5);
    deck.splice(0, 9);
    
    game.deck = deck;
    game.tableCard = game.deck.pop();
    game.tableCoins = 0;
    game.status = 'playing';
    
    game.turnIndex = Math.floor(Math.random() * players.length);
    game.turnStartTime = Date.now();
    
    let startCoins = 11;
    if (players.length === 6) startCoins = 9;
    if (players.length === 7) startCoins = 7;
    
    game.playerStates = {};
    players.forEach(p => {
        game.playerStates[p.id] = { coins: startCoins, cards: [] };
    });

    const firstPlayer = players[game.turnIndex];
    hostAnnounce(`เริ่มเกมแล้ว! เปิดไพ่ใบแรก เลข ${game.tableCard} เป็นรอบของ [PLAYER:${firstPlayer.id}]`, true);
}

function hostAnnounce(msg, assertive = false) {
    announce(msg, assertive);
    connections.forEach(conn => {
        if(conn.open) conn.send({ type: 'announce', message: msg, assertive: assertive });
    });
}

function processAction(player, action) {
    if (game.status !== 'playing') return;
    stopTimer();
    
    const pState = game.playerStates[player.id];
    
    if (action === 'pass') {
        if (pState.coins <= 0) return; 
        
        broadcastSound('skip');
        pState.coins -= 1;
        game.tableCoins += 1;
        
        hostAnnounce(`[PLAYER:${player.id}] เลือกผ่าน จ่าย 1 เหรียญ. เหรียญบนไพ่รวมเป็น ${game.tableCoins}`);
        
        game.turnIndex = (game.turnIndex + 1) % players.length;
        const nextPlayer = players[game.turnIndex];
        
        setTimeout(() => {
            game.turnStartTime = Date.now();
            hostAnnounce(`รอบของ[PLAYER:${nextPlayer.id}]`);
            broadcastState();
            checkBotTurn();
        }, 2000);

    } else if (action === 'take') {
        broadcastSound('jua');
        pState.cards.push(game.tableCard);
        pState.cards.sort((a,b) => a-b);
        const gainedCoins = game.tableCoins;
        pState.coins += gainedCoins;
        
        hostAnnounce(`[PLAYER:${player.id}] รับไพ่เลข ${game.tableCard} และรับเหรียญ ${gainedCoins} เหรียญ`, true);
        
        if (game.deck.length > 0) {
            game.tableCard = game.deck.pop();
            game.tableCoins = 0;
            
            setTimeout(() => {
                game.turnStartTime = Date.now();
                hostAnnounce(`เปิดไพ่ใบใหม่ เลข ${game.tableCard}. รอบของ[PLAYER:${player.id}] อีกครั้ง`);
                broadcastState();
                checkBotTurn();
            }, 2000);
        } else {
            game.tableCard = null; 
            game.tableCoins = 0;
            endGame();
        }
    }
}

function checkBotTurn() {
    if (!isHost || game.status !== 'playing') return;
    const currentPlayer = players[game.turnIndex];
    if (currentPlayer.isBot) {
        setTimeout(() => processBotTurn(), 1500); 
    }
}

function processBotTurn() {
    if (game.status !== 'playing') return;
    const bot = players[game.turnIndex];
    const state = game.playerStates[bot.id];
    
    if (state.coins === 0) { processAction(bot, 'take'); return; }

    const hasPrev = state.cards.includes(game.tableCard - 1);
    const hasNext = state.cards.includes(game.tableCard + 1);
    if (hasPrev || hasNext) { processAction(bot, 'take'); return; }

    const effectiveValue = game.tableCard - game.tableCoins;
    if (effectiveValue <= 5) { processAction(bot, 'take'); return; }

    processAction(bot, 'pass');
}

function updateVisualPlayers() {
    // ยกเลิกการสร้างตัวละครเรียงซ้ายขวาตามที่ได้รับมอบหมาย (ปิดการแสดงผล Visual Players Layer)
    const layer = document.getElementById('visual-players-layer');
    if (layer) {
        layer.innerHTML = '';
    }
    return;
}

function updateGameUI() {
    const isEn = currentLang === 'en';
    const currentPlayer = players[game.turnIndex];
    
    // --- ปรับสี Turn Banner ตามผู้เล่นปัจจุบัน ---
    const turnBanner = document.querySelector('.turn-banner');
    if (turnBanner && currentPlayer && currentPlayer.color) {
        turnBanner.setAttribute('data-color', currentPlayer.color.id);
    }
    
    document.getElementById('title-game').textContent = isEn
        ? `${getPronoun(currentPlayer)}'s Turn`
        : `รอบของ${getPronoun(currentPlayer)}`;

    if (game.status === 'playing' && currentPlayer && !currentPlayer.isBot) {
        startOrUpdateTimer();
    } else {
        stopTimer();
    }

    document.getElementById('deck-count').textContent = game.deck.length;
    document.getElementById('deck-info-container').setAttribute(
        'aria-label',
        isEn ? `${game.deck.length} cards left in deck` : `ไพ่ในกองเหลือ ${game.deck.length} ใบ`
    );
    
    const cardDisplay = document.getElementById('table-card-display');
    const coinDisplay = document.getElementById('table-coins-display');
    cardDisplay.textContent = game.tableCard !== null ? game.tableCard : "-";
    coinDisplay.textContent = game.tableCoins;
    
    if (game.tableCard === null) {
        cardDisplay.style.opacity = '0';
        coinDisplay.parentElement.style.opacity = '0';
    } else {
        cardDisplay.style.opacity = '1';
        coinDisplay.parentElement.style.opacity = '1';
    }

    const tableInfoStr = isEn
        ? `Table card is ${game.tableCard !== null ? game.tableCard : "-"}, ${game.tableCoins} coins on card`
        : `ไพ่บนโต๊ะคือเลข ${game.tableCard !== null ? game.tableCard : "-"} เหรียญสะสม ${game.tableCoins} เหรียญ`;
    document.getElementById('table-info-container').setAttribute('aria-label', tableInfoStr);

    const controls = document.getElementById('game-controls');
    const btnPass = document.getElementById('btn-pass');
    const btnTake = document.getElementById('btn-take');
    const btnRefreshAudio = document.getElementById('btn-refresh-audio');
    
    if (btnRefreshAudio) {
        btnRefreshAudio.style.display = (game.status === 'playing') ? 'block' : 'none';
    }

    const isMyTurn = currentPlayer.peerId === myPeerId;

    if (isMyTurn && !currentPlayer.isBot) {
        controls.style.display = 'flex';
        const myState = game.playerStates[currentPlayer.id];
        
        btnTake.disabled = false;
        btnTake.style.opacity = '1';
        
        if (myState.coins <= 0) {
            btnPass.disabled = true;
            btnPass.setAttribute(
                'aria-label',
                isEn ? 'Cannot pass, you have no coins left. You must take the card.' : 'ผ่านไม่ได้ คุณไม่มีเหรียญ ต้องรับไพ่เท่านั้น'
            );
        } else {
            btnPass.disabled = false;
            btnPass.setAttribute(
                'aria-label',
                isEn ? 'Pay 1 coin to pass' : 'จ่าย 1 เหรียญเพื่อผ่าน'
            );
        }
        
        setTimeout(() => {
            if(!btnPass.disabled) { btnPass.focus(); }
            else { btnTake.focus(); }
        }, 100);

        const currentStateStr = `${game.tableCard}-${game.tableCoins}`;
        if (myLastTurnState !== currentStateStr) {
            myLastTurnState = currentStateStr;
            if (turnAnnounceTimeout) clearTimeout(turnAnnounceTimeout);
            turnAnnounceTimeout = setTimeout(() => {
                if (game.status === 'playing' && players[game.turnIndex]?.peerId === myPeerId) {
                    announce(currentLang === 'en'
                        ? `It's your turn! Table card is ${game.tableCard !== null ? game.tableCard : "-"}, ${game.tableCoins} coins on card.`
                        : `ถึงรอบคุณแล้ว ไพ่บนโต๊ะคือเลข ${game.tableCard !== null ? game.tableCard : "-"} เหรียญสะสม ${game.tableCoins} เหรียญ`, true);
                }
            }, 1500);
        }
    } else {
        controls.style.display = 'none';
        myLastTurnState = null;
        if (turnAnnounceTimeout) {
            clearTimeout(turnAnnounceTimeout);
            turnAnnounceTimeout = null;
        }
    }

    const myStatusGroup = document.getElementById('my-status-group');
    const me = players.find(p => p.peerId === myPeerId);
    if (me) {
        const state = game.playerStates[me.id];
        const noCardsText = isEn ? 'No cards' : 'ไม่มีไพ่';
        const cardStr = state.cards.length > 0 ? state.cards.join(', ') : noCardsText;
        const isTurn = players[game.turnIndex].id === me.id;
        
        const ariaLabelText = isEn
            ? `${getPronoun(me)} has ${state.coins} coins. Cards in hand: ${cardStr}`
            : `${getPronoun(me)} มีเหรียญ ${state.coins} เหรียญ ไพ่ในมือ : ${cardStr}`;
        
        if (isTurn) {
            myStatusGroup.classList.add('active-turn');
        } else {
            myStatusGroup.classList.remove('active-turn');
        }

        let newCards = [];
        let coinGained = false;

        if (!window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
            if (visualStateMap.coins !== -1 && state.coins > visualStateMap.coins) {
                coinGained = true;
            }
            if (visualStateMap.cards !== null) {
                newCards = state.cards.filter(c => !visualStateMap.cards.includes(c));
            }
        }
        
        visualStateMap.coins = state.coins;
        visualStateMap.cards = [...state.cards];

        let cardsHTML = '';
        state.cards.forEach(c => {
            const animClass = newCards.includes(c) ? ' new-hand-card' : '';
            cardsHTML += `<div class="mini-card${animClass}">${c}</div>`;
        });
        if (state.cards.length === 0) {
            cardsHTML = `<div style="color:#aaa; font-size:14px; width:100%;">${noCardsText}</div>`;
        }
        
        const coinClass = coinGained ? 'my-status-coin coin-gain-anim' : 'my-status-coin';

        myStatusGroup.innerHTML = `
            <span class="sr-only">${ariaLabelText}</span>
            <div aria-hidden="true" class="my-status-layout">
                <div class="my-status-left">
                    <div class="${coinClass}">🪙 ${state.coins}</div>
                    <div class="hud-avatar" style="background-color: ${me.color.hex}; ${isTurn ? `box-shadow: 0 0 20px ${me.color.hex};` : ''}">
                        <span aria-hidden="true">^ᴗ^</span>
                    </div>
                </div>
                <div class="my-status-right">
                    <div style="font-size: 14px; font-weight: bold; color: ${me.color.hex}; margin-bottom: 8px;">${isTurn ? '⭐ ' : ''}${getPronoun(me)}</div>
                    <div class="my-cards-grid">
                        ${cardsHTML}
                    </div>
                </div>
            </div>
        `;
    }

    // Always call visual layers update
    updateVisualPlayers();
}

function disableActionButtons() {
    stopTimer();
    document.getElementById('btn-pass').disabled = true;
    document.getElementById('btn-take').disabled = true;
}

document.getElementById('btn-pass').addEventListener('click', () => {
    disableActionButtons();
    sendAction('pass');
});
document.getElementById('btn-take').addEventListener('click', () => {
    disableActionButtons();
    sendAction('take');
});

document.getElementById('btn-view-others').addEventListener('click', () => {
    const isEn = currentLang === 'en';
    const section = document.getElementById('all-players-section');
    const list = document.getElementById('all-players-list');
    list.innerHTML = '';
    
    document.getElementById('screen-game').classList.add('view-all-mode');
    
    players.forEach((p, index) => {
        const state = game.playerStates[p.id];
        const cardStr = state.cards.length > 0 ? state.cards.join(', ') : (isEn ? 'No cards' : 'ไม่มีไพ่');
        
        const isMe = (p.peerId === myPeerId && !p.isBot);
        const displayedCoins = isMe
            ? (isEn ? `${state.coins} coins` : `${state.coins} เหรียญ`)
            : (isEn ? '?? coins (Hidden)' : '?? เหรียญ (ซ่อน)');
        const ariaCoinsText = isMe
            ? `${state.coins}`
            : (isEn ? 'hidden' : 'ไม่สามารถดูได้');
        
        const div = document.createElement('div');
        div.className = `player-card`;
        
        const label = isEn
            ? `${getPronoun(p)} has ${ariaCoinsText} coins. Cards in hand: ${cardStr}`
            : `${getPronoun(p)} มีเหรียญ ${ariaCoinsText} เหรียญ ไพ่ในมือ : ${cardStr}`;
        
        div.setAttribute('tabindex', '0');
        div.setAttribute('role', 'listitem');
        
        div.innerHTML = `
            <span class="sr-only">${label}</span>
            <div aria-hidden="true" class="hud-content">
                <div class="hud-avatar" style="background-color: ${p.color.hex};">
                    <span aria-hidden="true">^ᴗ^</span>
                </div>
                <div class="hud-info">
                    <h3 style="color: ${p.color.hex}; margin-bottom: 0;">${getPronoun(p)}</h3>
                    <div class="hud-stats">
                        <span class="hud-coin" style="background: ${isMe ? 'linear-gradient(180deg, #ffca28, #ff8f00)' : '#555'}; color: ${isMe ? '#3e2723' : '#fff'};">🪙 ${displayedCoins}</span>
                    </div>
                    <div class="hud-cards">🃏 ${cardStr}</div>
                </div>
            </div>
        `;
        list.appendChild(div);
    });
    
    section.style.display = 'block';
    document.getElementById('title-all-players').focus();
    
    section.scrollIntoView({ behavior: 'smooth' });
});

document.getElementById('btn-close-all-players').addEventListener('click', () => {
    document.getElementById('all-players-section').style.display = 'none';
    document.getElementById('screen-game').classList.remove('view-all-mode');
    
    document.getElementById('btn-view-others').focus();
});

// --- ระบบคิดคะแนนและจบเกม ---
function calculateScore(cards, coins) {
    if (cards.length === 0) return -coins;
    let sorted = [...cards].sort((a,b) => a-b);
    let score = sorted[0];
    for (let i = 1; i < sorted.length; i++) {
        if (sorted[i] !== sorted[i-1] + 1) {
            score += sorted[i];
        }
    }
    return score - coins;
}

function runEndAnim() {
    const overlay = document.getElementById('anim-overlay');
    const textEl = document.getElementById('anim-text');
    overlay.style.display = 'flex';
    textEl.textContent = '';

    setTimeout(() => {
        textEl.textContent = currentLang === 'en' ? 'Game Over!' : 'จบเกม!';
        textEl.style.animation = 'none'; void textEl.offsetWidth; textEl.style.animation = 'bounceIn 0.8s cubic-bezier(0.175, 0.885, 0.32, 1.275)';
        announce(currentLang === 'en' ? 'Game Over' : 'จบเกม', true);
    }, 200);

    setTimeout(() => {
        textEl.textContent = currentLang === 'en' ? 'Calculating scores...' : 'กำลังทำการสรุปผล';
        textEl.style.animation = 'none'; void textEl.offsetWidth; textEl.style.animation = 'bounceIn 0.8s cubic-bezier(0.175, 0.885, 0.32, 1.275)';
        announce(currentLang === 'en' ? 'Calculating final scores' : 'กำลังทำการสรุปผล', true);
    }, 1200);

    setTimeout(() => {
        playEndGameSound();
    }, 1500);

    setTimeout(() => {
        overlay.style.display = 'none';
        textEl.textContent = '';
        document.getElementById('aria-polite').textContent = '';
        document.getElementById('aria-assertive').textContent = '';
        switchScreen('screen-result', 'title-result');
        updateResultUI();
    }, 2500);
}

function playEndGameSound() {
    if (!players || players.length === 0) return;
    const scores = players.map(p => game.playerStates[p.id]?.score ?? 999);
    const minScore = Math.min(...scores);
    const winners = players.filter(p => game.playerStates[p.id]?.score === minScore);
    const me = players.find(p => p.peerId === myPeerId);

    if (me && game.playerStates[me.id] && game.playerStates[me.id].score === minScore) {
        if (winners.length === 1) {
            playSound('win');
        } else {
            playSound('no');
        }
    } else {
        playSound('lost');
    }
}

function endGame() {
    stopTimer();
    game.status = 'ended';
    
    players.forEach(p => {
        const state = game.playerStates[p.id];
        state.score = calculateScore(state.cards, state.coins);
    });
    
    hostAnnounce("ไพ่หมดกองแล้ว จบเกม! กำลังสรุปผลคะแนน", true);
    connections.forEach(conn => {
        if(conn.open) conn.send({ type: 'triggerEndAnim', players, game });
    });
    runEndAnim();
}

function triggerConfetti() {
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    for(let i=0; i<40; i++) {
        let c = document.createElement('div');
        c.className = 'confetti';
        c.setAttribute('aria-hidden', 'true');
        c.style.left = Math.random() * 100 + 'vw';
        c.style.backgroundColor = ['#f44336', '#2196F3', '#4CAF50', '#FFEB3B', '#9C27B0', '#FF9800', '#E91E63'][Math.floor(Math.random()*7)];
        c.style.animationDuration = Math.random() * 2 + 3 + 's';
        c.style.animationDelay = Math.random() * 1 + 's';
        document.body.appendChild(c);
        setTimeout(() => { if (c.parentNode) c.parentNode.removeChild(c); }, 5000);
    }
}

function updateResultUI() {
    const isEn = currentLang === 'en';
    const list = document.getElementById('result-list');
    list.innerHTML = '';
    
    let results = players.map(p => {
        return {
            id: p.id,
            name: getPronoun(p),
            hex: p.color.hex,
            score: game.playerStates[p.id].score,
            cards: game.playerStates[p.id].cards,
            coins: game.playerStates[p.id].coins,
        };
    });
    
    results.sort((a, b) => a.score - b.score);
    let rankTextForSr = isEn ? "Score summary: " : "สรุปผลคะแนน: ";
    
    let me = players.find(p => p.peerId === myPeerId);
    
    results.forEach((r, index) => {
        const isWinner = index === 0 || r.score === results[0].score;
        const div = document.createElement('div');
        div.className = 'player-card' + (isWinner ? ' winner-card' : '');
        
        const rankHeading = isEn ? `Rank ${index + 1}: ${r.name}` : `อันดับ ${index + 1}: ${r.name}`;
        const scoreLine = isEn ? `Net Score: ${r.score}` : `คะแนนสุทธิ: ${r.score}`;
        const detailLine = isEn
            ? `Cards: ${r.cards.length > 0 ? r.cards.join(', ') : '-'} | Coins: ${r.coins}`
            : `ไพ่: ${r.cards.length > 0 ? r.cards.join(', ') : '-'} | เหรียญ: ${r.coins}`;

        div.innerHTML = `
            <div class="hud-content">
                <div class="hud-avatar" aria-hidden="true" style="background-color: ${r.hex}; border: 3px solid ${isWinner ? '#ffca28' : '#fff'};">
                    ${isWinner ? '🏆' : '<span aria-hidden="true">^ᴗ^</span>'}
                </div>
                <div class="hud-info">
                    <h3 style="color: ${isWinner ? '#ffca28' : r.hex}; margin-bottom: 5px;">${rankHeading}</h3>
                    <p style="font-size: 20px; color: white;"><strong>${scoreLine}</strong></p>
                    <p style="color: #bbb; font-size: 14px;">${detailLine}</p>
                </div>
            </div>
        `;
        list.appendChild(div);
        rankTextForSr += isEn
            ? `Rank ${index + 1} ${r.name} scored ${r.score} points. `
            : `อันดับ ${index + 1} ${r.name} ได้ ${r.score} คะแนน. `;
        
        if (isWinner && me && r.id === me.id) {
            triggerConfetti();
        }
    });

    announce(rankTextForSr);
}

// --- ระบบออกห้อง ---
function leaveRoom() {
    stopBGM();
    stopTimer();
    stopHeartbeat();
    resetJoinState();
    if (peer) peer.destroy(); 
    
    if (isHost && game.status === 'waiting' && currentRoomId) {
        remove(ref(db, `nothanks_rooms/${currentRoomId}`)); 
    }
    
    isHost = false;
    currentRoomId = null;
    players = [];
    connections = [];
    myLastTurnState = null;
    if (turnAnnounceTimeout) { clearTimeout(turnAnnounceTimeout); turnAnnounceTimeout = null; }
    visualStateMap.coins = -1;
    visualStateMap.cards = null;
    
    const turnBanner = document.querySelector('.turn-banner');
    if (turnBanner) turnBanner.removeAttribute('data-color');
    
    switchScreen('screen-main', 'title-main');
    initRoomListener(); 
    announce(currentLang === 'en' ? "Left the room. Returned to main menu." : "ออกจากห้อง กลับสู่เมนูหลักแล้ว");
}

document.getElementById('btn-leave-room').addEventListener('click', leaveRoom);
document.getElementById('btn-how-to-play').addEventListener('click', () => switchScreen('screen-rules', 'title-rules'));
document.getElementById('btn-back-main').addEventListener('click', () => switchScreen('screen-main', 'title-main'));

// เริ่มต้นการทำงานหน้าแรก
window.onload = () => {
    applyStaticLanguage();
    initRoomListener();
    setTimeout(() => {
        announce(currentLang === 'en'
            ? "Welcome to No Thanks! Focus is on the main menu."
            : "ยินดีต้อนรับเข้าสู่เกม No Thanks! เกมบริหารความเสี่ยง โฟกัสอยู่ที่เมนูหลักแล้ว");
        document.getElementById('title-main').focus();
    }, 500);
};

window.addEventListener('beforeunload', () => {
    stopBGM();
    stopTimer();
    stopHeartbeat();
    if (isHost && game.status === 'waiting' && currentRoomId) {
        remove(ref(db, `nothanks_rooms/${currentRoomId}`));
    }
});

// --- ระบบ Keyboard Shortcut ---
window.addEventListener('keydown', (event) => {
    if (game.status !== 'playing') return;

    if (event.altKey) {
        const key = event.key.toLowerCase();
        if (key === 's') {
            event.preventDefault();
            event.stopPropagation();
            const btnPass = document.getElementById('btn-pass');
            if (btnPass && !btnPass.disabled && btnPass.offsetParent !== null) {
                btnPass.click();
            }
        } else if (key === 'g') {
            event.preventDefault();
            event.stopPropagation();
            const btnTake = document.getElementById('btn-take');
            if (btnTake && !btnTake.disabled && btnTake.offsetParent !== null) {
                btnTake.click();
            }
        } else if (key === 'c') {
            event.preventDefault();
            event.stopPropagation();
            const me = players.find(p => p.peerId === myPeerId);
            if (me && game.playerStates[me.id]) {
                const state = game.playerStates[me.id];
                const isEn = currentLang === 'en';
                const cardStr = state.cards.length > 0 ? state.cards.join(', ') : (isEn ? 'No cards' : 'ไม่มีไพ่');
                const tableCardStr = game.tableCard !== null ? game.tableCard : "-";
                announce(isEn
                    ? `Table card ${tableCardStr}, ${game.tableCoins} coins on table. You have ${state.coins} coins. Cards in hand: ${cardStr}`
                    : `ไพ่บนโต๊ะเลข ${tableCardStr} เหรียญสะสม ${game.tableCoins} คุณมีเหรียญ ${state.coins} ไพ่ในมือ ${cardStr}`);
            }
        } else if (key === 'a') {
            event.preventDefault();
            event.stopPropagation();
            const btnViewOthers = document.getElementById('btn-view-others');
            if (btnViewOthers && btnViewOthers.offsetParent !== null) {
                btnViewOthers.click();
            }
        } else if (key === 'k') {
            event.preventDefault();
            event.stopPropagation();
            announce(currentLang === 'en'
                ? `${game.deck.length} cards left in deck`
                : `ไพ่ในกองเหลือ ${game.deck.length} ใบ`);
        } else if (key === 't') {
            event.preventDefault();
            event.stopPropagation();
            if (game.status === 'playing') {
                const currentPlayer = players[game.turnIndex];
                if (currentPlayer && !currentPlayer.isBot) {
                    const now = Date.now();
                    const elapsed = Math.floor((now - (game.turnStartTime || now)) / 1000);
                    const remaining = Math.max(0, 40 - elapsed);
                    announce(currentLang === 'en'
                        ? `${remaining} seconds remaining`
                        : `เหลือเวลาอีก ${remaining} วินาที`);
                }
            }
        }
    }
});

// --- ระบบรีเฟรชเสียง ---
document.getElementById('btn-refresh-audio').addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    
    try {
        if (window.AudioContext || window.webkitAudioContext) {
            if (audioCtx && audioCtx.state !== 'closed') {
                audioCtx.close().then(() => {
                    audioCtx = new (window.AudioContext || window.webkitAudioContext)();
                    audioCtx.resume();
                });
            } else {
                audioCtx = new (window.AudioContext || window.webkitAudioContext)();
                audioCtx.resume();
            }
        }
        
        announce(currentLang === 'en' ? "Audio system refreshed." : "รีเฟรชระบบเสียงเรียบร้อยแล้ว");
    } catch (err) {
        console.error("Audio refresh error:", err);
    }
});