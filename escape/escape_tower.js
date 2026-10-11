import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { getDatabase, ref, set, get, update, onValue, runTransaction, onDisconnect, remove } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-database.js";

// Your web app's Firebase configuration
const firebaseConfig = {
    apiKey: "AIzaSyDvcdgsyT5sDdYTYKIqetzNL9Be-MFC0l4",
    authDomain: "xo-game-134ec.firebaseapp.com",
    databaseURL: "https://xo-game-134ec-default-rtdb.asia-southeast1.firebasedatabase.app",
    projectId: "xo-game-134ec",
    storageBucket: "xo-game-134ec.firebasestorage.app",
    messagingSenderId: "318375224157",
    appId: "1:318375224157:web:667876ec774f9f03b77eb4"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

// --- ระบบประกาศผ่าน Screen Reader (Queue System) ---
const Announcer = {
    queue: [],
    isAnnouncing: false,
    add(msg) {
        this.queue.push(msg);
        this.playNext();
    },
    playNext() {
        if (this.isAnnouncing || this.queue.length === 0) return;
        this.isAnnouncing = true;
        const msg = this.queue.shift();
        const announcer = document.getElementById('announcer');
        if (announcer) {
            announcer.innerText = msg;
            setTimeout(() => {
                announcer.innerText = "";
                this.isAnnouncing = false;
                this.playNext();
            }, 2500); // เว้นระยะ 2.5 วินาที เพื่อให้ Screen Reader อ่านจบและไม่ขัดจังหวะนานเกินไป
        } else {
            this.isAnnouncing = false;
            this.playNext();
        }
    }
};

// --- สถาปัตยกรรมและอะแดปเตอร์ Multiplayer (Firebase + PeerJS) ---
const MultiplayerAdapter = {
    state: null,
    role: null, // 'host' หรือ 'client'
    roomId: null,
    peer: null,
    connections: [], // สำหรับ Host เก็บ connection ของ Client
    hostConnection: null, // สำหรับ Client เก็บ connection ไปยัง Host

    initLocal(stateObj) { 
        this.state = stateObj; 
    },

    async createRoom() {
        this.role = 'host';
        
        // รับหมายเลขห้องจาก Firebase Transaction แบบ Atomic
        const counterRef = ref(db, 'Escape/room_counter');
        let newRoomNumber = 1;
        
        await runTransaction(counterRef, (currentData) => {
            if (currentData === null) {
                return 1;
            } else {
                newRoomNumber = currentData + 1;
                return currentData + 1;
            }
        });

        // จัดรูปแบบห้อง Escape00001
        this.roomId = `Escape${String(newRoomNumber).padStart(5, '0')}`;

        // สร้าง Peer
        this.peer = new Peer(); 
        
        this.peer.on('open', (id) => {
            // บันทึกข้อมูลห้องลง Firebase
            const roomRef = ref(db, `Escape/rooms/${this.roomId}`);
            const roomData = {
                roomId: this.roomId,
                hostPeerId: id,
                status: 'waiting',
                playersCount: GameState.players.length
            };
            set(roomRef, roomData);
            
            // ตั้งค่าลบห้องอัตโนมัติเมื่อ Host หลุด
            onDisconnect(roomRef).remove();

            UI.announce(`สร้างห้องสำเร็จ หมายเลขห้อง: ${this.roomId}`);
            document.getElementById('room-creation').classList.remove('hidden');
        });

        this.peer.on('connection', (conn) => {
            conn.on('open', () => {
                this.connections.push(conn);
                this.sync(); // ส่ง State ปัจจุบันให้ Client ทันทีที่เชื่อมต่อ
            });

            conn.on('data', (data) => {
                if (data.type === 'ACTION') {
                    // ป้องกันการส่งคำสั่งซ้ำซ้อน
                    if (this.state && !this.state.isProcessingAction) {
                        this.state.isProcessingAction = true;
                        GameLogic.process(data.action, data.payload, data.clientId);
                        this.sync();
                        this.state.isProcessingAction = false;
                    }
                } else if (data.type === 'JOIN') {
                    // Client ขอ Join
                    if (GameState.isStarted) {
                        conn.send({ type: 'ERROR', message: 'เกมเริ่มไปแล้ว' });
                        return;
                    }
                    const newPlayer = GameLogic.createPlayer(data.playerName, false);
                    newPlayer.clientId = data.clientId; // แนบ ID อ้างอิง
                    GameState.players.push(newPlayer);
                    this.updateRoomFirebase();
                    this.sync();
                    UI.announce(`${data.playerName} เข้าร่วมห้องแล้ว`);
                    UI.updateRoomConfigUI();
                }
            });

            conn.on('close', () => {
                // เปลี่ยน Client เป็น Bot เมื่อหลุด
                const droppedPlayer = GameState.players.find(p => p.clientId === conn.peer);
                if (droppedPlayer) {
                    droppedPlayer.isBot = true;
                    UI.announce(`${droppedPlayer.name} หลุดการเชื่อมต่อ เปลี่ยนเป็นบอทอัตโนมัติ`);
                    this.sync();
                }
                this.connections = this.connections.filter(c => c.peer !== conn.peer);
            });
        });

        this.peer.on('error', (err) => {
            console.error(err);
            UI.announce("เกิดข้อผิดพลาดในการเชื่อมต่อเครือข่าย");
        });
    },

    joinRoom(roomId, playerName) {
        this.role = 'client';
        const roomRef = ref(db, `Escape/rooms/${roomId}`);
        
        get(roomRef).then((snapshot) => {
            if (snapshot.exists()) {
                const roomData = snapshot.val();
                if (roomData.status !== 'waiting') {
                    UI.announce("ไม่สามารถเข้าร่วมได้ เกมได้เริ่มไปแล้วหรือห้องปิดแล้ว");
                    return;
                }

                this.peer = new Peer();
                this.peer.on('open', (id) => {
                    this.hostConnection = this.peer.connect(roomData.hostPeerId);
                    
                    this.hostConnection.on('open', () => {
                        UI.localPlayerId = id;
                        this.hostConnection.send({ type: 'JOIN', playerName: playerName, clientId: id });
                        UI.showScreen('screen-lobby'); // ไปรอหน้า Lobby ของ Client
                        UI.announce(`กำลังเข้าร่วมห้อง ${roomId}...`);
                    });

                    this.hostConnection.on('data', (data) => {
                        if (data.type === 'SYNC') {
                            Object.assign(GameState, data.state);
                            if (GameState.isStarted && document.getElementById('screen-game').classList.contains('hidden')) {
                                UI.showScreen('screen-game');
                            }
                            UI.render();
                        } else if (data.type === 'ERROR') {
                            UI.announce(`ข้อผิดพลาด: ${data.message}`);
                        }
                    });

                    this.hostConnection.on('close', () => {
                        UI.announce("Host ปิดห้องหรือหลุดการเชื่อมต่อ ห้องถูกยุติแล้ว");
                        setTimeout(() => location.reload(), 3000);
                    });
                });
            } else {
                UI.announce("ไม่พบหมายเลขห้องนี้");
            }
        });
    },

    updateRoomFirebase() {
        if (this.role === 'host' && this.roomId) {
            const roomRef = ref(db, `Escape/rooms/${this.roomId}`);
            update(roomRef, {
                status: GameState.isStarted ? 'playing' : 'waiting',
                playersCount: GameState.players.length
            });
        }
    },

    closeRoom() {
        if (this.role === 'host' && this.roomId) {
            const roomRef = ref(db, `Escape/rooms/${this.roomId}`);
            remove(roomRef);
            this.connections.forEach(conn => conn.close());
        }
    },

    sync() {
        if (this.role === 'host') {
            const syncData = { type: 'SYNC', state: GameState };
            this.connections.forEach(conn => {
                if (conn.open) conn.send(syncData);
            });
            UI.render();
        }
    },

    dispatch(action, payload) {
        if (this.role === 'host') {
            GameLogic.process(action, payload, UI.localPlayerId);
            this.sync();
        } else if (this.role === 'client' && this.hostConnection && this.hostConnection.open) {
            this.hostConnection.send({ type: 'ACTION', action: action, payload: payload, clientId: UI.localPlayerId });
        }
    }
};

const GameState = {
    players: [],
    bots: [],
    floors: {},
    currentTurnIndex: 0,
    logs: [],
    isStarted: false,
    isProcessingAction: false
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
        
        rooms.sort(() => Math.random() - 0.5);
        
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
            clientId: null, // ผูกกับ Peer ID
            name: name,
            isBot: isBot,
            floor: 6,
            hp: 200,
            infection: 0,
            resting: false,
            inventory: {
                keys: 0,
                skeletonKeys: 0,
                vaccine: 0,
                food: 0,
                armor: 0,
                laserGuns: 0,
                ammo: 0
            },
            latestClueId: null,
            escaped: false,
            dead: false
        };
    },

    process(action, payload, requestClientId) {
        const player = GameState.players[GameState.currentTurnIndex];
        
        // ป้องกัน Client อื่นส่งคำสั่งข้ามเทิร์น
        if (!player.isBot && player.clientId !== requestClientId && UI.localPlayerId !== requestClientId) {
             // อนุญาตเฉพาะผู้เล่นที่ตรงกับเทิร์น หรือถ้า Host บังคับ
             return;
        }

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
                if (window.EscapeTowerClueBank) {
                    const clueObj = window.EscapeTowerClueBank.getClueById(room.clueId);
                    this.log(`${player.name} ได้รับคำใบ้: ${clueObj ? clueObj.question : 'ไม่พบข้อความ'}`);
                } else {
                    this.log(`${player.name} อ่านคำใบ้ปริศนาใหม่`);
                }
                break;
            case 'SECRET_PASSAGE':
                let targetFloor = Math.max(1, player.floor - 2);
                player.floor = targetFloor;
                player.latestClueId = null;
                this.log(`${player.name} พบทางลับ! ทิ้งดิ่งลงไปยังชั้น ${targetFloor}`);
                if (!GameState.floors[targetFloor]) this.generateFloor(targetFloor);
                break;
            case 'MYSTERY':
                this.log(`${player.name} พบห้องเหตุการณ์ลึกลับ ไม่พบอะไรผิดปกติ (ระบบเตรียมขยายผล)`);
                break;
        }
        
        if (player.hp < 200 && player.inventory.food > 0) {
            player.inventory.food--;
            player.hp = Math.min(200, player.hp + 30);
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
        
        if (player.infection === 0) {
            player.infection = 20;
            this.log(`${player.name} ถูกซอมบี้กัด! ติดเชื้อระดับ 1 (เสีย 20 HP/เทิร์น)`);
        } else {
            player.infection = 40;
            this.log(`${player.name} ถูกซอมบี้โจมตีซ้ำ! การติดเชื้อรุนแรงขึ้น (เสีย 40 HP/เทิร์น)`);
        }
        
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
            } else if (window.EscapeTowerClueBank && window.EscapeTowerClueBank.checkAnswer(player.latestClueId, payload.code)) {
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

            if (player.resting) {
                player.resting = false;
                player.hp = 100;
                this.log(`${player.name} ฟื้นตัวจากการพักฟื้น (HP กลับมาเป็น 100)`);
                continue;
            }

            if (player.infection > 0) {
                player.hp -= player.infection;
                this.log(`${player.name} เสีย ${player.infection} HP จากการติดเชื้อ (HP เหลือ ${player.hp})`);
                this.checkStatusLimits(player);
            }

            if (!player.dead && !player.resting) {
                break;
            }
            loopCount++;
        } while(loopCount < GameState.players.length);

        if (GameState.players.every(p => p.dead || p.escaped)) {
            this.log("เกมสิ้นสุดลงแล้ว!");
            return;
        }

        const activePlayer = GameState.players[GameState.currentTurnIndex];
        if (activePlayer) {
            this.log(`ถึงเทิร์นของ ${activePlayer.name} แล้ว`);
            if (activePlayer.isBot && MultiplayerAdapter.role === 'host') {
                setTimeout(() => BotController.playTurn(activePlayer), 1500);
            }
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
        if (!bot || bot.dead || bot.escaped) {
            GameLogic.nextTurn();
            return;
        }

        const floorData = GameState.floors[bot.floor];
        if (!floorData) {
            GameLogic.generateFloor(bot.floor);
        }
        
        const currentFloorRooms = GameState.floors[bot.floor].rooms;
        let availableRooms = currentFloorRooms.map((r, i) => ({room: r, index: i}))
                             .filter(item => item.room.exploredBy.length === 0 || item.room.type === 'CLUE');
        
        if (bot.inventory.skeletonKeys > 0) {
            MultiplayerAdapter.dispatch('USE_DOOR', { method: 'SKELETON_KEY' });
            return;
        } else if (bot.inventory.keys > 0) {
            MultiplayerAdapter.dispatch('USE_DOOR', { method: 'KEY' });
            return;
        }

        if (bot.latestClueId && window.EscapeTowerClueBank) {
            const clue = window.EscapeTowerClueBank.getClueById(bot.latestClueId);
            if (clue && clue.answer) {
                MultiplayerAdapter.dispatch('USE_DOOR', { method: 'CODE', code: clue.answer });
                return;
            }
        }

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
        const playerNameInput = document.getElementById('player-name');
        
        // กฎข้อ 5: การกด Enter เพื่อยืนยันชื่อ
        if (playerNameInput) {
            playerNameInput.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    this.handleLogin();
                }
            });
        }
        document.getElementById('btn-start').addEventListener('click', this.handleLogin.bind(this));
        
        document.getElementById('btn-create-room').addEventListener('click', () => {
            document.getElementById('btn-create-room').disabled = true;
            MultiplayerAdapter.createRoom();
        });

        // ดึงรายการห้องจาก Firebase
        const roomListRef = ref(db, 'Escape/rooms');
        onValue(roomListRef, (snapshot) => {
            this.renderRoomList(snapshot.val());
        });
        
        document.getElementById('btn-add-bot').addEventListener('click', this.handleAddBot.bind(this));
        document.getElementById('btn-remove-bot').addEventListener('click', this.handleRemoveBot.bind(this));
        document.getElementById('btn-start-game').addEventListener('click', this.handleStartGame.bind(this));
        
        const btnViewAll = document.getElementById('btn-view-all-players');
        if (btnViewAll) {
            btnViewAll.addEventListener('click', this.openAllPlayersStatusModal.bind(this));
        }
        const btnCloseAll = document.getElementById('btn-close-all-players-status');
        if (btnCloseAll) {
            btnCloseAll.addEventListener('click', this.closeAllPlayersStatusModal.bind(this));
        }

        const btnMainDoor = document.getElementById('btn-main-door');
        if (btnMainDoor) {
            btnMainDoor.addEventListener('click', this.toggleDoorControls.bind(this));
        }

        document.getElementById('btn-use-key').addEventListener('click', () => {
            MultiplayerAdapter.dispatch('USE_DOOR', { method: 'KEY' });
        });
        document.getElementById('btn-use-skeleton-key').addEventListener('click', () => {
            MultiplayerAdapter.dispatch('USE_DOOR', { method: 'SKELETON_KEY' });
        });
        document.getElementById('btn-open-keypad').addEventListener('click', this.openKeypad.bind(this));
        
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
            const modalKeypad = document.getElementById('modal-keypad');
            if (modalKeypad && typeof modalKeypad.close === 'function') {
                modalKeypad.close();
            }
            const btnOpenKeypad = document.getElementById('btn-open-keypad');
            if (btnOpenKeypad) btnOpenKeypad.focus();
        });
        document.getElementById('btn-keypad-submit').addEventListener('click', () => {
            MultiplayerAdapter.dispatch('USE_DOOR', { method: 'CODE', code: this.currentCode });
            const modalKeypad = document.getElementById('modal-keypad');
            if (modalKeypad && typeof modalKeypad.close === 'function') {
                modalKeypad.close();
            }
        });

        // ตรวจสอบการโหลดไฟล์คลังคำใบ้
        if (!window.EscapeTowerClueBank) {
            this.announce("ข้อผิดพลาด: ไม่พบไฟล์ escape_tower_clues.js เกมไม่สามารถทำงานได้");
        }
        
        // จัดการเมื่อปิดหน้าต่าง
        window.addEventListener("beforeunload", () => {
            MultiplayerAdapter.closeRoom();
        });
    },

    renderRoomList(rooms) {
        const listContainer = document.getElementById('room-list') || document.createElement('div');
        if (!document.getElementById('room-list')) {
            listContainer.id = 'room-list';
            document.getElementById('screen-login').appendChild(listContainer);
        }
        listContainer.innerHTML = '<h3>ห้องที่สามารถเข้าร่วมได้</h3>';
        
        let hasWaitingRooms = false;
        if (rooms) {
            Object.values(rooms).forEach(room => {
                // กฎข้อ 3: ห้องที่เริ่มเกมแล้ว ต้องหายไปจากรายการห้องสาธารณะ
                if (room.status === 'waiting') {
                    hasWaitingRooms = true;
                    const roomBtn = document.createElement('button');
                    roomBtn.innerText = `เข้าร่วมห้อง ${room.roomId} (ผู้เล่น: ${room.playersCount})`;
                    roomBtn.className = 'room-btn';
                    roomBtn.onclick = () => {
                        const nameInput = document.getElementById('player-name').value.trim();
                        if (!nameInput) {
                            this.announce("กรุณากรอกชื่อผู้เล่นก่อนเข้าร่วมห้อง");
                            return;
                        }
                        MultiplayerAdapter.joinRoom(room.roomId, nameInput);
                    };
                    listContainer.appendChild(roomBtn);
                }
            });
        }
        
        if (!hasWaitingRooms) {
            listContainer.innerHTML += '<p>ไม่มีห้องว่างในขณะนี้ กรุณาสร้างห้องใหม่</p>';
        }
    },

    showScreen(id) {
        document.querySelectorAll('.screen').forEach(s => {
            s.classList.remove('active');
            s.classList.add('hidden');
        });
        const target = document.getElementById(id);
        if (target) {
            target.classList.remove('hidden');
            target.classList.add('active');
            target.focus(); 
        }
    },

    announce(text) {
        Announcer.add(text);
        
        const logArea = document.getElementById('event-log');
        if (logArea) {
            const p = document.createElement('p');
            p.innerText = text;
            logArea.prepend(p);
        }
    },

    handleLogin() {
        const nameInput = document.getElementById('player-name').value.trim();
        if (!nameInput) {
            this.announce("กรุณากรอกชื่อผู้เล่น");
            return;
        }
        // ป้องกันการกดสร้างผู้เล่นซ้ำในกรณีที่ไม่ใช่ Client ที่กำลัง Join
        if (!MultiplayerAdapter.role) {
            const newPlayer = GameLogic.createPlayer(nameInput, false);
            newPlayer.clientId = this.localPlayerId = 'local_' + Date.now();
            GameState.players = [newPlayer]; // รีเซ็ตเพื่อความแน่ใจ
            document.getElementById('display-player-name').innerText = nameInput;
            this.showScreen('screen-lobby');
            this.announce("เข้าสู่ล็อบบี้แล้ว คุณคือ Host");
            document.getElementById('btn-start').disabled = true;
        }
    },

    handleAddBot() {
        if (MultiplayerAdapter.role !== 'host') return;
        if (GameState.players.length < 4) {
            GameState.players.push(GameLogic.createPlayer(`Bot-${GameState.players.length}`, true));
            this.updateRoomConfigUI();
            MultiplayerAdapter.updateRoomFirebase();
            this.announce("เพิ่มบอทแล้ว");
        }
    },

    handleRemoveBot() {
        if (MultiplayerAdapter.role !== 'host') return;
        const lastPlayer = GameState.players[GameState.players.length - 1];
        if (lastPlayer && lastPlayer.isBot) {
            GameState.players.pop();
            this.updateRoomConfigUI();
            MultiplayerAdapter.updateRoomFirebase();
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
        if (MultiplayerAdapter.role !== 'host') {
            this.announce("เฉพาะ Host เท่านั้นที่เริ่มเกมได้");
            return;
        }
        // กฎข้อ 6: ต้องมีตัวละครอย่างน้อย 2 ตัวจึงจะเริ่มเกมได้
        if (GameState.players.length < 2) {
            this.announce("ไม่สามารถเริ่มเกมได้ ต้องมีผู้เล่นหรือบอทรวมกันอย่างน้อย 2 ตัวละคร");
            return;
        }
        
        if (GameState.isStarted) return;
        GameState.isStarted = true;
        
        MultiplayerAdapter.initLocal(GameState);
        MultiplayerAdapter.updateRoomFirebase(); // เปลี่ยนสถานะเป็น playing
        
        GameLogic.generateFloor(6);
        this.showScreen('screen-game');
        GameLogic.log("เริ่มเกม! ทุกคนอยู่ชั้น 6");
        MultiplayerAdapter.sync();
    },

    toggleDoorControls() {
        const doorControls = document.getElementById('door-controls');
        const btnMainDoor = document.getElementById('btn-main-door');
        if (doorControls && btnMainDoor) {
            const isHidden = doorControls.classList.contains('hidden');
            if (isHidden) {
                doorControls.classList.remove('hidden');
                btnMainDoor.setAttribute('aria-expanded', 'true');
                const firstOption = doorControls.querySelector('button:not(:disabled)');
                if (firstOption) firstOption.focus();
            } else {
                doorControls.classList.add('hidden');
                btnMainDoor.setAttribute('aria-expanded', 'false');
                btnMainDoor.focus();
            }
        }
    },

    openAllPlayersStatusModal() {
        const modal = document.getElementById('modal-all-players-status');
        if (modal) {
            this.renderAllPlayersList();
            if (typeof modal.showModal === 'function') {
                modal.showModal();
            } else {
                modal.classList.remove('hidden');
            }
            const closeBtn = document.getElementById('btn-close-all-players-status');
            if (closeBtn) closeBtn.focus();
        }
    },

    closeAllPlayersStatusModal() {
        const modal = document.getElementById('modal-all-players-status');
        if (modal) {
            if (typeof modal.close === 'function') {
                modal.close();
            } else {
                modal.classList.add('hidden');
            }
            const btnViewAll = document.getElementById('btn-view-all-players');
            if (btnViewAll && btnViewAll.offsetWidth > 0 && btnViewAll.offsetHeight > 0) {
                btnViewAll.focus();
            }
        }
    },

    renderAllPlayersList() {
        const listContainer = document.getElementById('all-players-list');
        if (!listContainer) return;

        listContainer.innerHTML = '';
        GameState.players.forEach((p, idx) => {
            const card = document.createElement('div');
            card.className = `player-status-card ${idx === GameState.currentTurnIndex ? 'active-turn' : ''}`;
            card.innerHTML = `
                <strong>${p.name}</strong> ${p.isBot ? '(บอท)' : ''} (ชั้น ${p.floor})<br>
                HP: ${p.hp} ${p.infection > 0 ? `[ติดเชื้อ -\${p.infection}]` : ''} ${p.resting ? '[พักฟื้น]' : ''} ${p.escaped ? '[รอดชีวิต]' : ''} ${p.dead ? '[เสียชีวิต]' : ''}<br>
                ไอเท็ม: กุญแจ(${p.inventory.keys}) ผี(${p.inventory.skeletonKeys})<br>
                ปืนเลเซอร์: ${p.inventory.laserGuns} (กระสุน ${p.inventory.ammo})<br>
                เกราะ: ${p.inventory.armor} | วัคซีน: ${p.inventory.vaccine} | อาหาร: ${p.inventory.food}
            `;
            listContainer.appendChild(card);
        });
    },

    render() {
        // หาตัว Local Player (จาก clientId ของ Peer หรือ id อ้างอิง)
        const localPlayer = GameState.players.find(p => p.clientId === this.localPlayerId) || GameState.players.find(p => p.clientId === null) || GameState.players[0];
        const activePlayer = GameState.players[GameState.currentTurnIndex];
        const isMyTurn = localPlayer && !localPlayer.dead && !localPlayer.escaped && activePlayer && (localPlayer.id === activePlayer.id);

        if (localPlayer) {
            document.getElementById('current-floor-display').innerText = localPlayer.floor;
        }

        if (activePlayer) {
            document.getElementById('turn-indicator').innerText = activePlayer.name;
        }

        const localStatusContainer = document.getElementById('local-player-status');
        if (localStatusContainer && localPlayer) {
            localStatusContainer.innerHTML = `
                <strong>คุณ: ${localPlayer.name}</strong> (ชั้น ${localPlayer.floor})<br>
                HP: ${localPlayer.hp} ${localPlayer.infection > 0 ? `[ติดเชื้อ -\${localPlayer.infection}]` : ''} ${localPlayer.resting ? '[พักฟื้น]' : ''}<br>
                ไอเท็ม: กุญแจ(${localPlayer.inventory.keys}) ผี(${localPlayer.inventory.skeletonKeys}) | ปืนเลเซอร์: ${localPlayer.inventory.laserGuns} (กระสุน ${localPlayer.inventory.ammo})<br>
                เกราะ: ${localPlayer.inventory.armor} | วัคซีน: ${localPlayer.inventory.vaccine} | อาหาร: ${localPlayer.inventory.food}
            `;
        }

        const statusPanel = document.getElementById('status-panel');
        if (statusPanel) {
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
        }

        const modalAll = document.getElementById('modal-all-players-status');
        if (modalAll && modalAll.open) {
            this.renderAllPlayersList();
        }

        const board = document.getElementById('game-board');
        if (board && localPlayer) {
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
        }

        const btnMainDoor = document.getElementById('btn-main-door');
        if (btnMainDoor) {
            btnMainDoor.disabled = !isMyTurn;
        }

        const btnUseKey = document.getElementById('btn-use-key');
        if (btnUseKey && localPlayer) {
            btnUseKey.disabled = !isMyTurn || localPlayer.inventory.keys === 0;
            btnUseKey.innerText = `ใช้กุญแจธรรมดา (${localPlayer.inventory.keys})`;
        }

        const btnUseSkeletonKey = document.getElementById('btn-use-skeleton-key');
        if (btnUseSkeletonKey && localPlayer) {
            btnUseSkeletonKey.disabled = !isMyTurn || localPlayer.inventory.skeletonKeys === 0;
            btnUseSkeletonKey.innerText = `ใช้กุญแจผี (${localPlayer.inventory.skeletonKeys})`;
        }

        const btnOpenKeypad = document.getElementById('btn-open-keypad');
        if (btnOpenKeypad) {
            btnOpenKeypad.disabled = !isMyTurn;
            btnOpenKeypad.innerText = "ใส่รหัสคำใบ้";
        }
    },

    showClueModal(room, isMyTurn, roomIndex) {
        const clue = window.EscapeTowerClueBank ? window.EscapeTowerClueBank.getClueById(room.clueId) : null;
        if (isMyTurn && !room.exploredBy.includes(this.localPlayerId)) {
            MultiplayerAdapter.dispatch('EXPLORE', { roomIndex: roomIndex });
        }
        if (clue) {
            // ใช้ Announcer สำหรับผู้ใช้ Screen Reader
            this.announce(`ทบทวนปริศนา: ${clue.question}`);
            
            // ใช้โครงสร้าง UI เดิมหรือแจ้งเตือนที่เข้าถึงได้ หากมีโครงสร้าง Modal ให้เปลี่ยนมาใช้ แต่ถ้าไม่มีให้ใช้ข้อความบนจอแทน alert()
            const clueArea = document.getElementById('clue-display-area');
            if(clueArea) {
                clueArea.innerText = `ปริศนา: ${clue.question}`;
            } else {
                alert(`ปริศนา:\n${clue.question}`);
            }
        }
    },

    openKeypad() {
        const localPlayer = GameState.players.find(p => p.clientId === this.localPlayerId) || GameState.players[0];
        if (!localPlayer || !localPlayer.latestClueId) {
            this.announce("คุณยังไม่มีคำใบ้ของชั้นนี้");
            return;
        }
        const clue = window.EscapeTowerClueBank ? window.EscapeTowerClueBank.getClueById(localPlayer.latestClueId) : null;
        if (clue) {
            document.getElementById('clue-display-area').innerText = `ปริศนาของคุณ: ${clue.question}`;
            this.currentCode = "";
            this.updateKeypadDisplay();
            const modalKeypad = document.getElementById('modal-keypad');
            if (modalKeypad && typeof modalKeypad.showModal === 'function') {
                modalKeypad.showModal();
            }
        }
    },

    updateKeypadDisplay() {
        document.getElementById('keypad-input').innerText = this.currentCode || "_";
    }
};

window.addEventListener('DOMContentLoaded', () => UI.init());