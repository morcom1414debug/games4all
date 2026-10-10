"use strict";

// --- สถาปัตยกรรมและอะแดปเตอร์ ---
const MultiplayerAdapter = {
    // โครงสร้างเตรียมพร้อมเชื่อม Firebase
    state: null,
    initLocal(stateObj) { this.state = stateObj; },
    sync() { /* ในอนาคตเชื่อม Firebase Set/Update */ UI.render(); },
    dispatch(action, payload) { GameLogic.process(action, payload); this.sync(); }
};

const GameState = {
    players: [],
    bots: [],
    floors: {}, // บันทึกข้อมูลแต่ละชั้น
    currentTurnIndex: 0,
    logs: [],
    isStarted: false
};

const ROOM_TYPES = [
    { type: 'KEY', count: 1, name: 'กุญแจธรรมดา' },
    { type: 'SKELETON_KEY', count: 1, name: 'กุญแจผี' },
    { type: 'SECRET_PASSAGE', count: 1, name: 'ทางลับ' },
    { type: 'CLUE', count: 3, name: 'ห้องคำใบ้' },
    { type: 'ZOMBIE', count: 4, name: 'ซอมบี้' },
    { type: 'VACCINE', count: 3, name: 'วัคซีน' },
    { type: 'FOOD', count: 3, name: 'อาหาร' },
    { type: 'ARMOR', count: 1, name: 'เกราะสวรรค์' },
    { type: 'MYSTERY', count: 1, name: 'เหตุการณ์ลึกลับ' },
    { type: 'LASER_GUN', count: 1, name: 'ปืนเลเซอร์' },
    { type: 'BATTERY', count: 1, name: 'แบตเตอรี่ปืนเลเซอร์' }
];

// --- การจัดการ Logic หลัก ---
const GameLogic = {
    generateFloor(floorNum) {
        if(GameState.floors[floorNum]) return;
        let rooms = [];
        ROOM_TYPES.forEach(rt => {
            for(let i=0; i<rt.count; i++) rooms.push({ type: rt.type, id: `F${floorNum}_R${rooms.length}`, exploredBy: [] });
        });
        // สุ่มห้อง
        rooms.sort(() => Math.random() - 0.5);
        
        // กำหนดคำใบ้
        let clues = [];
        if (window.EscapeTowerClueBank) {
            clues = window.EscapeTowerClueBank.selectCluesForGame(floorNum);
        } else {
            console.error("EscapeTowerClueBank ไม่พบไฟล์คลังปริศนา");
        }

        let clueIndex = 0;
        rooms.forEach(r => {
            if(r.type === 'CLUE' && clues[clueIndex]) {
                r.clueId = clues[clueIndex].id;
                clueIndex++;
            }
        });

        GameState.floors[floorNum] = { rooms: rooms };
    },

    createPlayer(name, isBot = false) {
        return {
            id: 'P_' + Date.now() + Math.random(),
            name: name,
            isBot: isBot,
            floor: 6,
            hp: 200,
            infection: 0, // 0, 20, 40
            resting: false,
            inventory: {
                keys: 0,
                skeletonKeys: 0,
                vaccine: 0,
                food: 0, // max 3
                armor: 0, // 0 or 1
                laserGuns: 0,
                ammo: 0 // max 2
            },
            latestClueId: null,
            escaped: false,
            dead: false
        };
    },

    process(action, payload) {
        const player = GameState.players[GameState.currentTurnIndex];
        if(player.dead || player.escaped) {
            this.nextTurn();
            return;
        }

        if (action === 'EXPLORE') {
            this.handleExplore(player, payload.roomIndex);
        } else if (action === 'USE_DOOR') {
            this.handleDoor(player, payload);
        }
    },

    handleExplore(player, roomIndex) {
        const floorData = GameState.floors[player.floor];
        const room = floorData.rooms[roomIndex];

        if (room.exploredBy.length > 0 && room.type !== 'CLUE') {
            this.log(`ห้องนี้ถูกสำรวจไปแล้ว`);
            return;
        }
        
        if (!room.exploredBy.includes(player.id)) {
            room.exploredBy.push(player.id);
        }

        this.resolveRoomEvent(player, room);
        this.checkStatusLimits(player);
        this.nextTurn();
    },

    resolveRoomEvent(player, room) {
        switch(room.type) {
            case 'KEY': 
                player.inventory.keys++; 
                this.log(`${player.name} พบกุญแจธรรมดา 1 ดอก`); 
                break;
            case 'SKELETON_KEY': 
                player.inventory.skeletonKeys++; 
                this.log(`${player.name} พบกุญแจผี 1 ดอก`); 
                break;
            case 'VACCINE': 
                if (player.infection > 0) {
                    player.infection = 0;
                    this.log(`${player.name} พบวัคซีนและใช้รักษาการติดเชื้อทันที`);
                } else {
                    player.inventory.vaccine++;
                    this.log(`${player.name} พบวัคซีน 1 หลอด เก็บไว้ในคลัง`);
                }
                break;
            case 'FOOD':
                if (player.inventory.food < 3) {
                    player.inventory.food++;
                    this.log(`${player.name} พบอาหาร 1 ชิ้น เก็บเข้าคลัง`);
                } else {
                    this.log(`${player.name} พบอาหาร แต่คลังอาหารเต็มแล้ว (สูงสุด 3 ชิ้น)`);
                }
                break;
            case 'ARMOR':
                player.inventory.armor = 1;
                this.log(`${player.name} พบเกราะสวรรค์! จะป้องกันการโจมตีจากซอมบี้ได้ 1 ครั้ง`);
                break;
            case 'LASER_GUN':
                player.inventory.laserGuns++;
                if (player.inventory.ammo < 2) {
                    player.inventory.ammo++;
                    this.log(`${player.name} พบปืนเลเซอร์ ได้รับกระสุน 1 นัด`);
                } else {
                    this.log(`${player.name} พบปืนเลเซอร์ แต่กระสุนเต็มแล้ว`);
                }
                break;
            case 'BATTERY':
                if (player.inventory.ammo < 2) {
                    player.inventory.ammo++;
                    this.log(`${player.name} พบแบตเตอรี่ปืนเลเซอร์ ได้รับกระสุนเพิ่ม 1 นัด`);
                } else {
                    this.log(`${player.name} พบแบตเตอรี่ปืนเลเซอร์ แต่กระสุนเต็มแล้ว`);
                }
                break;
            case 'ZOMBIE':
                this.handleZombieEncounter(player);
                break;
            case 'CLUE':
                player.latestClueId = room.clueId;
                this.log(`${player.name} อ่านคำใบ้ปริศนาใหม่`);
                break;
            case 'SECRET_PASSAGE':
                let targetFloor = Math.max(1, player.floor - 2);
                player.floor = targetFloor;
                player.latestClueId = null; // ล้างคำใบ้
                this.log(`${player.name} พบทางลับ! ทิ้งดิ่งลงไปยังชั้น ${targetFloor}`);
                if (!GameState.floors[targetFloor]) this.generateFloor(targetFloor);
                break;
            case 'MYSTERY':
                this.log(`${player.name} พบห้องเหตุการณ์ลึกลับ ไม่พบอะไรผิดปกติ (ระบบเตรียมขยายผล)`);
                break;
        }
        
        // ฟื้นฟูอัตโนมัติหาก HP ลด
        if (player.hp < 200 && player.inventory.food > 0) {
            player.inventory.food--;
            player.hp = Math.min(200, player.hp + 30); // สมมติฟื้น 30
            this.log(`${player.name} กินอาหาร ฟื้นฟู HP เป็น ${player.hp}`);
        }
    },

    handleZombieEncounter(player) {
        if (player.inventory.ammo > 0 && player.inventory.laserGuns > 0) {
            player.inventory.ammo--;
            this.log(`${player.name} เผชิญหน้าซอมบี้! ใช้ปืนเลเซอร์ยิงสกัดสำเร็จ (กระสุนเหลือ ${player.inventory.ammo})`);
            return;
        }
        if (player.inventory.armor > 0) {
            player.inventory.armor = 0;
            this.log(`${player.name} เผชิญหน้าซอมบี้! เกราะสวรรค์แตก แต่ป้องกันการโจมตีได้`);
            return;
        }
        
        // ติดเชื้อ
        if (player.infection === 0) {
            player.infection = 20;
            this.log(`${player.name} ถูกซอมบี้กัด! ติดเชื้อระดับ 1 (เสีย 20 HP/เทิร์น)`);
        } else {
            player.infection = 40;
            this.log(`${player.name} ถูกซอมบี้โจมตีซ้ำ! การติดเชื้อรุนแรงขึ้น (เสีย 40 HP/เทิร์น)`);
        }
        
        // ใช้วัคซีนอัตโนมัติถ้ามี
        if (player.inventory.vaccine > 0) {
            player.inventory.vaccine--;
            player.infection = 0;
            this.log(`${player.name} ใช้วัคซีนฉุกเฉิน รักษาการติดเชื้อทันที!`);
        }
    },

    handleDoor(player, payload) {
        let success = false;
        if (payload.method === 'KEY' && player.inventory.keys > 0) {
            player.inventory.keys--;
            success = true;
            this.log(`${player.name} ใช้กุญแจธรรมดาไขประตูสำเร็จ`);
        } else if (payload.method === 'SKELETON_KEY' && player.inventory.skeletonKeys > 0) {
            player.inventory.skeletonKeys--;
            success = true;
            this.log(`${player.name} ใช้กุญแจผีเปิดประตูสำเร็จ`);
        } else if (payload.method === 'CODE') {
            if (!player.latestClueId) {
                this.log(`${player.name} ไม่มีคำใบ้ของชั้นนี้ จึงกดรหัสไม่ได้`);
            } else if (window.EscapeTowerClueBank.checkAnswer(player.latestClueId, payload.code)) {
                success = true;
                this.log(`${player.name} กรอกรหัสถูกต้อง! ประตูเปิดออก`);
            } else {
                this.log(`${player.name} กรอกรหัสผิด! ประตูยังคงล็อกอยู่`);
                this.nextTurn();
                return;
            }
        }

        if (success) {
            player.floor--;
            player.latestClueId = null;
            if (player.floor < 1) {
                player.escaped = true;
                this.log(`>>> ${player.name} หนีออกจากตึกมรณะสำเร็จ! <<<`);
            } else {
                this.log(`${player.name} ลงบันไดไปสู่ชั้น ${player.floor}`);
                if (!GameState.floors[player.floor]) this.generateFloor(player.floor);
            }
            this.nextTurn();
        }
    },

    nextTurn() {
        let loopCount = 0;
        do {
            GameState.currentTurnIndex = (GameState.currentTurnIndex + 1) % GameState.players.length;
            const player = GameState.players[GameState.currentTurnIndex];
            
            if (player.dead || player.escaped) continue;

            // ตรวจสอบพักฟื้น
            if (player.resting) {
                player.resting = false;
                player.hp = 100;
                this.log(`${player.name} ฟื้นตัวจากการพักฟื้น (HP กลับมาเป็น 100)`);
                continue; // ข้ามเทิร์น
            }

            // รับดาเมจติดเชื้อ
            if (player.infection > 0) {
                player.hp -= player.infection;
                this.log(`${player.name} เสีย ${player.infection} HP จากการติดเชื้อ (HP เหลือ ${player.hp})`);
                this.checkStatusLimits(player);
            }

            if (!player.dead && !player.resting) {
                break; // พบผู้เล่นที่พร้อมเล่น
            }
            loopCount++;
        } while(loopCount < GameState.players.length);

        if (GameState.players.every(p => p.dead || p.escaped)) {
            this.log("เกมสิ้นสุดลงแล้ว!");
            return;
        }

        const activePlayer = GameState.players[GameState.currentTurnIndex];
        if (activePlayer && activePlayer.isBot) {
            setTimeout(() => BotController.playTurn(activePlayer), 1000);
        }
    },

    checkStatusLimits(player) {
        player.hp = Math.min(200, Math.max(0, player.hp));
        if (player.hp === 0) {
            player.dead = true;
            this.log(`${player.name} เสียชีวิต...`);
        } else if (player.hp <= 60 && !player.resting) {
            player.resting = true;
            this.log(`${player.name} บาดเจ็บหนัก ต้องเข้าสู่สถานะพักฟื้นในเทิร์นถัดไป`);
        }
    },

    log(msg) {
        GameState.logs.push(msg);
        UI.announce(msg);
    }
};

const BotController = {
    playTurn(bot) {
        const floorData = GameState.floors[bot.floor];
        let availableRooms = floorData.rooms.map((r, i) => ({room: r, index: i}))
                             .filter(item => item.room.exploredBy.length === 0 || item.room.type === 'CLUE');
        
        // บอทไขกุญแจถ้ามี
        if (bot.inventory.skeletonKeys > 0) {
            MultiplayerAdapter.dispatch('USE_DOOR', { method: 'SKELETON_KEY' });
            return;
        } else if (bot.inventory.keys > 0) {
            MultiplayerAdapter.dispatch('USE_DOOR', { method: 'KEY' });
            return;
        }

        // บอทรู้คำตอบถ้ามี Clue (จำลอง)
        if (bot.latestClueId) {
            const clue = window.EscapeTowerClueBank.getClueById(bot.latestClueId);
            MultiplayerAdapter.dispatch('USE_DOOR', { method: 'CODE', code: clue.answer });
            return;
        }

        // ถ้าเปิดประตูไม่ได้ ให้สุ่มห้องที่เข้าได้
        if (availableRooms.length > 0) {
            const choice = availableRooms[Math.floor(Math.random() * availableRooms.length)];
            MultiplayerAdapter.dispatch('EXPLORE', { roomIndex: choice.index });
        } else {
            GameLogic.log(`${bot.name} ไม่เหลือทางไป...`);
            GameLogic.nextTurn();
        }
    }
};

// --- การจัดการ UI และ DOM ---
const UI = {
    currentCode: "",
    localPlayerId: null,

    init() {
        document.getElementById('btn-start').addEventListener('click', this.handleLogin.bind(this));
        document.getElementById('btn-create-room').addEventListener('click', () => {
            document.getElementById('btn-create-room').disabled = true;
            document.getElementById('room-creation').classList.remove('hidden');
            UI.announce("สร้างห้องสำเร็จ รอเพิ่มบอทหรือเริ่มเกม");
        });
        document.getElementById('btn-add-bot').addEventListener('click', this.handleAddBot.bind(this));
        document.getElementById('btn-remove-bot').addEventListener('click', this.handleRemoveBot.bind(this));
        document.getElementById('btn-start-game').addEventListener('click', this.handleStartGame.bind(this));
        
        // แผงควบคุมประตู
        document.getElementById('btn-use-key').addEventListener('click', () => {
            MultiplayerAdapter.dispatch('USE_DOOR', { method: 'KEY' });
        });
        document.getElementById('btn-use-skeleton-key').addEventListener('click', () => {
            MultiplayerAdapter.dispatch('USE_DOOR', { method: 'SKELETON_KEY' });
        });
        document.getElementById('btn-open-keypad').addEventListener('click', this.openKeypad.bind(this));
        
        // Keypad
        document.querySelectorAll('.key-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                this.currentCode += e.target.dataset.val;
                this.updateKeypadDisplay();
            });
        });
        document.getElementById('btn-keypad-clear').addEventListener('click', () => {
            this.currentCode = "";
            this.updateKeypadDisplay();
        });
        document.getElementById('btn-keypad-del').addEventListener('click', () => {
            this.currentCode = this.currentCode.slice(0, -1);
            this.updateKeypadDisplay();
        });
        document.getElementById('btn-keypad-cancel').addEventListener('click', () => {
            document.getElementById('modal-keypad').close();
        });
        document.getElementById('btn-keypad-submit').addEventListener('click', () => {
            MultiplayerAdapter.dispatch('USE_DOOR', { method: 'CODE', code: this.currentCode });
            document.getElementById('modal-keypad').close();
        });

        if (!window.EscapeTowerClueBank) {
            alert("ข้อผิดพลาด: ไม่พบไฟล์ escape_tower_clues.js เกมไม่สามารถทำงานได้");
        }
    },

    showScreen(id) {
        document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
        document.getElementById(id).classList.add('active');
        document.getElementById(id).focus(); // Accessibility
    },

    announce(text) {
        document.getElementById('announcer').innerText = text;
        const logArea = document.getElementById('event-log');
        const p = document.createElement('p');
        p.innerText = text;
        logArea.prepend(p);
    },

    handleLogin() {
        const nameInput = document.getElementById('player-name').value.trim();
        if (!nameInput) {
            this.announce("กรุณากรอกชื่อผู้เล่น");
            return;
        }
        GameState.players.push(GameLogic.createPlayer(nameInput, false));
        this.localPlayerId = GameState.players[0].id;
        document.getElementById('display-player-name').innerText = nameInput;
        this.showScreen('screen-lobby');
        this.announce("เข้าสู่ล็อบบี้แล้ว");
    },

    handleAddBot() {
        if (GameState.players.length < 4) {
            GameState.players.push(GameLogic.createPlayer(`Bot-${GameState.players.length}`, true));
            this.updateRoomConfigUI();
            this.announce("เพิ่มบอทแล้ว");
        }
    },

    handleRemoveBot() {
        const lastPlayer = GameState.players[GameState.players.length - 1];
        if (lastPlayer && lastPlayer.isBot) {
            GameState.players.pop();
            this.updateRoomConfigUI();
            this.announce("ลดบอทแล้ว");
        }
    },

    updateRoomConfigUI() {
        document.getElementById('count-players').innerText = GameState.players.length;
        const botCount = GameState.players.filter(p => p.isBot).length;
        document.getElementById('count-bots').innerText = botCount;
        document.getElementById('btn-add-bot').disabled = GameState.players.length >= 4;
        document.getElementById('btn-remove-bot').disabled = botCount <= 0;
    },

    handleStartGame() {
        if (GameState.isStarted) return;
        GameState.isStarted = true;
        MultiplayerAdapter.initLocal(GameState);
        GameLogic.generateFloor(6);
        this.showScreen('screen-game');
        GameLogic.log("เริ่มเกม! ทุกคนอยู่ชั้น 6");
        MultiplayerAdapter.sync();
    },

    render() {
        const localPlayer = GameState.players.find(p => p.id === this.localPlayerId);
        const activePlayer = GameState.players[GameState.currentTurnIndex];
        const isMyTurn = localPlayer && !localPlayer.dead && !localPlayer.escaped && (localPlayer.id === activePlayer.id);

        document.getElementById('current-floor-display').innerText = localPlayer.floor;
        document.getElementById('turn-indicator').innerText = `เทิร์นของ: ${activePlayer.name}`;

        // Render Status Panel
        const statusPanel = document.getElementById('status-panel');
        statusPanel.innerHTML = '';
        GameState.players.forEach((p, idx) => {
            const div = document.createElement('div');
            div.className = `player-status-card ${idx === GameState.currentTurnIndex ? 'active-turn' : ''}`;
            div.innerHTML = `
                <strong>${p.name}</strong> (ชั้น ${p.floor})<br>
                HP: ${p.hp} ${p.infection > 0 ? `[ติดเชื้อ -\${p.infection}]` : ''} ${p.resting ? '[พักฟื้น]' : ''}<br>
                ไอเท็ม: กุญแจ(${p.inventory.keys}) ผี(${p.inventory.skeletonKeys})<br>
                ปืนเลเซอร์: ${p.inventory.laserGuns} (กระสุน ${p.inventory.ammo})<br>
                เกราะ: ${p.inventory.armor} | วัคซีน: ${p.inventory.vaccine} | อาหาร: ${p.inventory.food}
            `;
            statusPanel.appendChild(div);
        });

        // Render Board
        const board = document.getElementById('game-board');
        board.innerHTML = '';
        const floorData = GameState.floors[localPlayer.floor];
        if (floorData) {
            floorData.rooms.forEach((room, index) => {
                const btn = document.createElement('button');
                btn.className = 'room-btn';
                
                const isExplored = room.exploredBy.length > 0;
                if (isExplored && room.type !== 'CLUE') {
                    btn.classList.add('explored');
                    btn.disabled = true;
                    btn.innerText = room.name;
                    btn.setAttribute('aria-label', `ห้อง ${index + 1} สำรวจแล้ว (${room.name})`);
                } else if (room.type === 'CLUE' && room.exploredBy.includes(localPlayer.id)) {
                    btn.classList.add('clue-room');
                    btn.innerText = "คำใบ้ (อ่านแล้ว)";
                    btn.setAttribute('aria-label', `ห้องคำใบ้ ${index + 1} อ่านแล้ว กดเพื่ออ่านซ้ำ`);
                    btn.onclick = () => this.showClueModal(room, isMyTurn, index);
                } else {
                    btn.innerText = `ห้อง ${index + 1}`;
                    btn.disabled = !isMyTurn;
                    btn.onclick = () => MultiplayerAdapter.dispatch('EXPLORE', { roomIndex: index });
                }
                board.appendChild(btn);
            });
        }

        // Render Door Controls
        document.getElementById('btn-use-key').disabled = !isMyTurn || localPlayer.inventory.keys === 0;
        document.getElementById('btn-use-skeleton-key').disabled = !isMyTurn || localPlayer.inventory.skeletonKeys === 0;
        document.getElementById('btn-open-keypad').disabled = !isMyTurn;
    },

    showClueModal(room, isMyTurn, roomIndex) {
        const clue = window.EscapeTowerClueBank.getClueById(room.clueId);
        if(isMyTurn && !room.exploredBy.includes(this.localPlayerId)) {
            MultiplayerAdapter.dispatch('EXPLORE', { roomIndex: roomIndex });
        }
        alert(`ปริศนา:\n${clue.question}`); // เพื่อความเรียบง่ายของ UI ใช้ alert สำหรับอ่านซ้ำ
    },

    openKeypad() {
        const localPlayer = GameState.players.find(p => p.id === this.localPlayerId);
        if (!localPlayer.latestClueId) {
            this.announce("คุณยังไม่มีคำใบ้ของชั้นนี้");
            return;
        }
        const clue = window.EscapeTowerClueBank.getClueById(localPlayer.latestClueId);
        document.getElementById('clue-display-area').innerText = `ปริศนาของคุณ: ${clue.question}`;
        this.currentCode = "";
        this.updateKeypadDisplay();
        document.getElementById('modal-keypad').showModal();
    },

    updateKeypadDisplay() {
        document.getElementById('keypad-input').innerText = this.currentCode || "_";
    }
};

window.addEventListener('DOMContentLoaded', () => UI.init());