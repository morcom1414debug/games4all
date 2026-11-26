// === Crossroad: Audio-Only & Visual Cartoon Road-Crossing Game ===
/*
  Architecture:
    - LaneManager: Handles lanes, car spawning, and per-lane logic
    - CarSpawner: Controls timing and spawning of cars in lanes with vehicle types & random colors
    - CollisionDetector: Checks for danger zone collisions (synced with vehicle size)
    - AudioMixer: Handles directional mixing, volume, panning, overlapping
    - VisualRenderer: 3rd-person 2.5D cartoon renderer on <canvas> (100% hidden from Screen Readers)
    - CrossroadGame: Controls game loop, focus mode, ARIA feedback, pause/resume, restart
*/

// -------- Bilingual Localization (Standard Pattern) --------
const LANG_STORAGE_KEY = 'crossroad_lang';

const I18N = {
  th: {
    pageTitle: 'ผู้พิทักษ์ไร้สายตา: ฝ่าวิกฤตถนนเดือด — เกมเสียงสำหรับคนตาบอด',
    metaDescription: 'เกมข้ามถนนด้วยเสียงสำหรับผู้พิการทางสายตา',
    mainAria: 'เกมข้ามถนน',
    h1: 'ผู้พิทักษ์ไร้สายตา: ฝ่าวิกฤตถนนเดือด: เกมเสียงสำหรับคนตาบอด',
    backBtn: 'ย้อนกลับ',
    langBtnText: '🌐 EN',
    langBtnAria: 'สลับภาษาเป็นภาษาอังกฤษ (Switch to English)',
    langChangedSR: 'เปลี่ยนภาษาเป็นภาษาไทยแล้ว',
    controlsAria: 'ปุ่มควบคุมเกม',
    startBtn: 'เริ่มเล่นเกม',
    resumeBtn: 'เล่นต่อ',
    crossBtn: 'ข้ามถนน',
    crossBtnAriaInit: 'ข้ามถนน (แตะ 3 ครั้ง / กดลูกศรขึ้น 3 ครั้ง)',
    restartBtn: 'เริ่มใหม่',
    restartBtnAria: 'เริ่มเกมใหม่',
    instIntro: 'อ่านคำแนะนำเพื่อฟังวิธีเล่น กดลูกศรขึ้น 3 ครั้ง หรือกดปุ่ม "ข้ามถนน" 3 ครั้ง เพื่อข้ามเลน',
    pausedSR: 'หยุดเกมชั่วคราว ออกจากโหมดโฟกัสแล้ว กดปุ่ม เล่นต่อ เพื่อดำเนินการต่อ กด R เพื่อเริ่มใหม่ หรือกด Escape เพื่อกลับหน้าหลัก',
    returnedMainSR: 'กลับมาหน้าหลักของเกมแล้ว กดปุ่ม เริ่มเล่นเกม เพื่อเริ่มต้นใหม่',
    pressMore: (n) => `กดอีก ${n} ครั้งเพื่อข้ามเลน`,
    lanePassed: (lane, roads, crossed, score) => `ข้ามเลนที่ ${lane} สำเร็จ! คุณข้ามถนนได้ ${roads} เส้น รวม ${crossed} เลน คะแนนรวม ${score} คะแนน`,
    stageReach: (stage, lanes) => `ถึงด่านที่ ${stage} แล้ว ตอนนี้มี ${lanes} เลน`,
    stageReady: 'เตรียมตัวสำหรับด่านถัดไป!',
    statusWin: 'เกมจบแล้ว! คุณชนะ',
    statusLoss: 'คุณถูกรถชน ลองใหม่หรือฟังคำแนะนำอีกครั้ง',
    statusPlaying: (s, totalS, l, totalL) => `ด่านที่ ${s}/${totalS}, เลนที่ ${l}/${totalL}`,
    crossAriaPlaying: (l, totalL) => `ข้ามเลน: กด 3 ครั้งเพื่อข้ามเลนที่ ${l} จาก ${totalL}`,
    summary: (roads, crossed, totalL, score) => `คุณข้ามถนนได้ ${roads} เส้น รวม ${crossed} เลน (จากทั้งหมด ${totalL} เลน) คะแนนรวม ${score} คะแนน`,
    srLoss: (sum) => `ชน! คุณถูกรถชน ${sum} กด R หรือปุ่มเริ่มใหม่เพื่อเล่นอีกครั้ง หรือกด Escape เพื่อกลับหน้าหลัก`,
    srWin: (sum) => `ยินดีด้วย! คุณข้ามถนนได้สำเร็จและชนะแล้ว ${sum} กด R เพื่อเล่นอีกครั้ง หรือกด Escape เพื่อกลับหน้าหลัก`,
    restarting: 'กำลังเริ่มเกมใหม่...',
    // Visual Canvas Labels
    canvasTitle: 'ผู้พิทักษ์ไร้สายตา',
    canvasSubtitle: 'กด "เริ่มเล่นเกม" เพื่อช่วยผู้คนข้ามถนน',
    canvasStageHud: (s, totalS) => `ด่าน ${s}/${totalS}`,
    canvasLaneHud: (l, totalL) => `เลน ${l}/${totalL}`,
    canvasScoreHud: (score) => `${score} คะแนน`,
    canvasSafeZone: 'ปลอดภัย! ข้ามครบแล้ว',
    canvasBonk: 'โอ๊ย! ชนเข้าแล้ว!',
    canvasStageClear: (s) => `ผ่านด่านที่ ${s} แล้ว!`,
    canvasVictory: 'ภารกิจสำเร็จ! ปลอดภัยทุกคน!',
    storyHTML: `<h2>เรื่องราว</h2>
      <p>คุณคืออาสาสมัครผู้พิการทางสายตา ที่มีความกล้าหาญและความตั้งใจจะช่วยเหลือผู้คนข้ามถนนที่วุ่นวายและเต็มไปด้วยรถยนต์ในเมืองไทย ทุกการข้ามถนนคือความท้าทายที่ต้องใช้ประสาทสัมผัสทางเสียงและความระมัดระวังอย่างสูงสุด เพื่อให้ทุกคนข้ามไปถึงฝั่งอย่างปลอดภัย</p>
      <h2>วิธีเล่น</h2>
      <ul>
        <li>กดปุ่ม <b>เริ่มเล่นเกม</b> เพื่อเริ่มต้นภารกิจ</li>
        <li>ฟังเสียงรถที่กำลังแล่นผ่านแต่ละเลน และใช้เสียงเป็นตัวช่วยในการตัดสินใจ</li>
        <li>ข้ามแต่ละเลนโดย <b>กดปุ่ม "ข้ามถนน" 3 ครั้ง</b> หรือ <b>กดลูกศรขึ้น 3 ครั้ง</b> เพื่อข้ามเลน</li>
        <li>ระวัง! หากข้ามในช่วงที่มีรถอยู่ในจุดอันตราย จะเกิดอุบัติเหตุทันที</li>
        <li>ข้ามถนนให้ครบทุกเลนเพื่อพิชิตภารกิจและช่วยเหลือทุกคนอย่างปลอดภัย</li>
        <li>หากต้องการเริ่มใหม่ กดปุ่ม <b>เริ่มใหม่</b> หรือกดตัว <b>R</b> บนแป้นพิมพ์</li>
      </ul>`
  },
  en: {
    pageTitle: 'Blind Guardian: Rush Hour Crossing — Audio Game for the Blind',
    metaDescription: 'Audio-based road-crossing game for visually impaired players',
    mainAria: 'Road Crossing Game',
    h1: 'Blind Guardian: Rush Hour Crossing — Audio Game for the Blind',
    backBtn: 'Back',
    langBtnText: '🌐 ไทย',
    langBtnAria: 'Switch language to Thai (เปลี่ยนเป็นภาษาไทย)',
    langChangedSR: 'Language switched to English',
    controlsAria: 'Game controls',
    startBtn: 'Start Game',
    resumeBtn: 'Resume',
    crossBtn: 'Cross Road',
    crossBtnAriaInit: 'Cross Road (Tap 3 times / Press Up Arrow 3 times)',
    restartBtn: 'Restart',
    restartBtnAria: 'Restart Game',
    instIntro: 'Read the instructions to learn how to play. Press Up Arrow 3 times or press the "Cross Road" button 3 times to cross a lane.',
    pausedSR: 'Game paused. Exited focus mode. Press Resume to continue, press R to restart, or press Escape to return to the main menu.',
    returnedMainSR: 'Returned to the main menu. Press Start Game to begin.',
    pressMore: (n) => `Press ${n} more time${n > 1 ? 's' : ''} to cross the lane`,
    lanePassed: (lane, roads, crossed, score) => `Lane ${lane} crossed! You crossed ${roads} road(s), ${crossed} lane(s) total. Total score: ${score} points`,
    stageReach: (stage, lanes) => `Reached Stage ${stage}. Now there are ${lanes} lanes.`,
    stageReady: 'Get ready for the next stage!',
    statusWin: 'Game Over! You Win!',
    statusLoss: 'You were hit by a car. Try again or read the instructions again.',
    statusPlaying: (s, totalS, l, totalL) => `Stage ${s}/${totalS}, Lane ${l}/${totalL}`,
    crossAriaPlaying: (l, totalL) => `Cross lane: Press 3 times to cross lane ${l} of ${totalL}`,
    summary: (roads, crossed, totalL, score) => `You crossed ${roads} road(s), ${crossed} lane(s) (out of ${totalL} total lanes). Total score: ${score} points`,
    srLoss: (sum) => `Crash! You were hit by a car. ${sum}. Press R or the Restart button to play again, or press Escape to return to the main menu.`,
    srWin: (sum) => `Congratulations! You crossed the roads safely and won! ${sum}. Press R to play again, or press Escape to return to the main menu.`,
    restarting: 'Restarting game...',
    // Visual Canvas Labels
    canvasTitle: 'Blind Guardian',
    canvasSubtitle: 'Press "Start Game" to help people cross',
    canvasStageHud: (s, totalS) => `Stage ${s}/${totalS}`,
    canvasLaneHud: (l, totalL) => `Lane ${l}/${totalL}`,
    canvasScoreHud: (score) => `${score} PTS`,
    canvasSafeZone: 'SAFE! ALL LANES CROSSED',
    canvasBonk: 'BONK! WATCH OUT!',
    canvasStageClear: (s) => `STAGE ${s} CLEAR!`,
    canvasVictory: 'MISSION COMPLETE!',
    storyHTML: `<h2>Story</h2>
      <p>You are a courageous visually impaired volunteer determined to help people cross busy, traffic-filled streets. Every road crossing is a challenge requiring acute auditory awareness and utmost care to guide everyone safely to the other side.</p>
      <h2>How to Play</h2>
      <ul>
        <li>Press the <b>Start Game</b> button to begin your mission.</li>
        <li>Listen carefully to the cars passing in each lane and use sound cues to decide when to move.</li>
        <li>Cross each lane by <b>pressing the "Cross Road" button 3 times</b> or <b>pressing the Up Arrow key 3 times</b>.</li>
        <li>Watch out! Crossing while a car is in the danger zone will cause an immediate crash.</li>
        <li>Cross all lanes across every stage to complete the mission and save everyone safely.</li>
        <li>To play again, press the <b>Restart</b> button or press <b>R</b> on your keyboard.</li>
      </ul>`
  }
};

function detectInitialLanguage() {
  try {
    const saved = localStorage.getItem(LANG_STORAGE_KEY);
    if (saved === 'th' || saved === 'en') return saved;
  } catch (e) {}
  const browserLang = (navigator.language || navigator.userLanguage || 'en').toLowerCase();
  return browserLang.startsWith('th') ? 'th' : 'en';
}

let currentLang = detectInitialLanguage();

function t(key) {
  return I18N[currentLang][key];
}

function applyLanguageUI() {
  document.documentElement.lang = currentLang;
  document.title = t('pageTitle');

  const metaDesc = document.querySelector('meta[name="description"]');
  if (metaDesc) metaDesc.setAttribute('content', t('metaDescription'));

  const mainEl = document.getElementById('main');
  if (mainEl) mainEl.setAttribute('aria-label', t('mainAria'));

  const titleEl = document.getElementById('title');
  if (titleEl) titleEl.innerText = t('h1');

  const backBtn = document.getElementById('btn-back');
  if (backBtn) backBtn.innerText = t('backBtn');

  const langBtn = document.getElementById('btn-lang-toggle');
  if (langBtn) {
    langBtn.innerText = t('langBtnText');
    langBtn.setAttribute('aria-label', t('langBtnAria'));
  }

  const controlsEl = document.getElementById('controls');
  if (controlsEl) controlsEl.setAttribute('aria-label', t('controlsAria'));

  const startBtn = document.getElementById('btn-start-game');
  if (startBtn) {
    const btnText = (window.game && window.game.isPaused) ? t('resumeBtn') : t('startBtn');
    startBtn.innerText = btnText;
    startBtn.setAttribute('aria-label', btnText);
  }

  const crossBtn = document.getElementById('btn-cross');
  if (crossBtn) {
    crossBtn.innerText = t('crossBtn');
    if (window.game && window.game.gameActive) {
      crossBtn.setAttribute('aria-label', t('crossAriaPlaying')(window.game.currentLaneIdx + 1, STAGES[window.game.stageIdx].lanes));
    } else {
      crossBtn.setAttribute('aria-label', t('crossBtnAriaInit'));
    }
  }

  const restartBtn = document.getElementById('btn-restart');
  if (restartBtn) {
    restartBtn.innerText = t('restartBtn');
    restartBtn.setAttribute('aria-label', t('restartBtnAria'));
  }

  const storyEl = document.getElementById('story');
  if (storyEl) storyEl.innerHTML = t('storyHTML');

  const instRegion = document.getElementById('instructions-region');
  if (instRegion && instRegion.innerText.trim() !== '') {
    instRegion.innerText = t('instIntro');
  }

  if (window.game && (window.game.lossState || window.game.winState)) {
    window.game.updateEndResultUI(false);
  }
}

function toggleLanguage() {
  currentLang = (currentLang === 'th') ? 'en' : 'th';
  try {
    localStorage.setItem(LANG_STORAGE_KEY, currentLang);
  } catch (e) {}
  applyLanguageUI();
  updateSRFeedback(t('langChangedSR'));
}

// -------- Vehicle Types & Cartoon Color Palettes (1 - 6) --------
const VEHICLE_CATALOG = [
  { id: 1, type: 'sedan',        nameTh: 'รถยนต์ 4 ที่นั่ง',       width: 72,  height: 36, dangerStart: 0.30, dangerEnd: 0.60 },
  { id: 2, type: 'suv7',         nameTh: 'รถยนต์ 7 ที่นั่ง',       width: 86,  height: 44, dangerStart: 0.28, dangerEnd: 0.62 },
  { id: 3, type: 'van',          nameTh: 'รถตู้ 10-12 ที่นั่ง',     width: 102, height: 48, dangerStart: 0.26, dangerEnd: 0.64 },
  { id: 4, type: 'moto_female',  nameTh: 'มอเตอร์ไซค์ (ผู้หญิงขับ)', width: 46,  height: 38, dangerStart: 0.35, dangerEnd: 0.55 },
  { id: 5, type: 'moto_male',    nameTh: 'มอเตอร์ไซค์ (ผู้ชายขับ)',  width: 48,  height: 38, dangerStart: 0.35, dangerEnd: 0.55 },
  { id: 6, type: 'bus',          nameTh: 'รถโดยสารประจำทาง',      width: 134, height: 54, dangerStart: 0.22, dangerEnd: 0.68 }
];

const VEHICLE_PALETTE = [
  { name: 'แดง',       body: '#FF3B30', shade: '#C92A21', accent: '#FFD60A' },
  { name: 'น้ำเงิน',    body: '#007AFF', shade: '#0056B3', accent: '#FFFFFF' },
  { name: 'เขียว',      body: '#34C759', shade: '#248A3D', accent: '#FFF3B0' },
  { name: 'เหลือง',     body: '#FFCC00', shade: '#D4A000', accent: '#1C1C1E' },
  { name: 'ทอง',        body: '#D4AF37', shade: '#9E7F1E', accent: '#FFF8DC' },
  { name: 'เงิน',       body: '#D1D5DB', shade: '#9CA3AF', accent: '#3B82F6' },
  { name: 'ขาว',        body: '#F9FAFB', shade: '#D1D5DB', accent: '#EF4444' },
  { name: 'ฟ้า',        body: '#38BDF8', shade: '#0284C7', accent: '#FFFFFF' },
  { name: 'ชมพู',       body: '#FF6B9E', shade: '#D6336C', accent: '#FFF0F5' },
  { name: 'ส้ม',        body: '#FF9500', shade: '#C96D00', accent: '#FFFFFF' },
  { name: 'ม่วง',       body: '#AF52DE', shade: '#7B2CBF', accent: '#F3E8FF' },
  { name: 'เขียวมินต์',   body: '#00C7BE', shade: '#008E87', accent: '#FFFFFF' },
  { name: 'ครีม',       body: '#FDE68A', shade: '#D97706', accent: '#991B1B' },
  { name: 'น้ำตาลช็อกโกแลต', body: '#8D5B4C', shade: '#5C3A30', accent: '#FDE68A' }
];

function createVehicleVisualMeta(slotIndex) {
  const spec = VEHICLE_CATALOG[slotIndex] || VEHICLE_CATALOG[0];
  const color = VEHICLE_PALETTE[Math.floor(Math.random() * VEHICLE_PALETTE.length)];
  // สำหรับรถเมล์ (id === 6) สุ่มว่าเป็นรถเมล์ร้อน (พัดลม) หรือรถเมล์ปรับอากาศ
  const isAirConBus = spec.id === 6 ? Math.random() > 0.45 : false;
  return {
    ...spec,
    color,
    isAirConBus,
    wheelPhase: Math.random() * Math.PI * 2
  };
}

// -------- Audio Asset Paths --------
const AUDIO = {
  car: {
    l: ['audio/car-l1.mp3', 'audio/car-l2.mp3', 'audio/car-l3.mp3', 'audio/car-l4.mp3', 'audio/car-l5.mp3', 'audio/car-l6.mp3'],
    r: ['audio/car-r1.mp3', 'audio/car-r2.mp3', 'audio/car-r3.mp3', 'audio/car-r4.mp3', 'audio/car-r5.mp3', 'audio/car-r6.mp3']
  },
  crash: 'audio/car-crash.mp3',
  lanePass: 'audio/lane-pass.mp3',
  win: 'audio/win-sound.mp3',
  instructions: 'audio/instructions.mp3'
};

// ----------- Game Config -----------
const STAGES = [
  { lanes: 2, carDensity: 0.7 },
  { lanes: 4, carDensity: 1.0 },
  { lanes: 6, carDensity: 1.3 },
  { lanes: 8, carDensity: 1.6 },
  { lanes: 10, carDensity: 1.9 },
  { lanes: 12, carDensity: 2.2 },
  { lanes: 14, carDensity: 2.5 },
  { lanes: 16, carDensity: 2.8 }
];

const CONTROLS_REQUIRED_CLICKS = 3;

// ----------- Audio Helper -----------
class AudioPlayer {
  constructor() {
    this.ctx = window.AudioContext ? new AudioContext() : null;
    this.buffers = {};
  }

  async loadBuffer(url) {
    if (this.buffers[url]) return this.buffers[url];
    const res = await fetch(url);
    const arrBuf = await res.arrayBuffer();
    const buffer = await this.ctx.decodeAudioData(arrBuf);
    this.buffers[url] = buffer;
    return buffer;
  }

  async play(url, { volume = 1, onEnded = null } = {}) {
    if (!this.ctx) {
      const audio = new Audio(url);
      audio.volume = volume;
      audio.addEventListener('ended', () => {
        if (onEnded) onEnded();
      });
      audio.play();
      return { audio, startedAt: Date.now() / 1000, duration: audio.duration || 3, url };
    }
    if (this.ctx.state === 'suspended') {
      try { await this.ctx.resume(); } catch (e) {}
    }
    const buffer = await this.loadBuffer(url);
    const source = this.ctx.createBufferSource();
    source.buffer = buffer;

    let gain = this.ctx.createGain();
    gain.gain.value = volume;

    source.connect(gain).connect(this.ctx.destination);

    const startedAt = this.ctx.currentTime;
    source.start();

    source.onended = () => {
      if (onEnded) onEnded();
    };

    return {
      source,
      gainNode: gain,
      startedAt: startedAt,
      duration: buffer.duration,
      url
    };
  }
}

// -------- LaneManager & CarSpawner --------
class LaneManager {
  constructor(stageIdx) {
    this.stageIdx = stageIdx;
    this.setUpLanes();
  }

  setUpLanes() {
    this.laneConfigs = [];
    const laneCount = STAGES[this.stageIdx].lanes;
    const halfLane = Math.floor(laneCount / 2);
    for (let i = 0; i < laneCount; ++i) {
      let dir = i < halfLane ? 'r' : 'l';
      this.laneConfigs.push({ direction: dir });
    }
  }

  startSpawners(audioMixer) {
    this.activeCarsPerLane = [];
    this.spawners = [];
    const density = STAGES[this.stageIdx].carDensity;
    for (let i = 0; i < this.laneConfigs.length; ++i) {
      let spawner = new CarSpawner(this.laneConfigs[i].direction, audioMixer, density, (carObj) => {
        if (!this.activeCarsPerLane[i]) this.activeCarsPerLane[i] = [];
        this.activeCarsPerLane[i].push(carObj);
        carObj.endCallback = () => {
          this.activeCarsPerLane[i] = this.activeCarsPerLane[i].filter(obj => obj !== carObj);
        };
      });
      spawner.laneIdx = i;
      spawner.start();
      this.spawners.push(spawner);
      this.activeCarsPerLane.push([]);
    }
  }

  stopSpawnerForLane(laneIdx) {
    if (this.spawners && this.spawners[laneIdx]) {
      this.spawners[laneIdx].stop();
    }
    if (this.activeCarsPerLane && this.activeCarsPerLane[laneIdx]) {
      this.activeCarsPerLane[laneIdx].forEach(car => {
        if (car.source) { try { car.source.stop(); } catch (e) { } }
        if (car.audio) { try { car.audio.pause(); } catch (e) { } }
      });
      this.activeCarsPerLane[laneIdx] = [];
    }
  }

  stopSpawners() {
    this.spawners?.forEach(spawner => spawner.stop());
  }

  getActiveCarsInLane(laneIdx) {
    return this.activeCarsPerLane[laneIdx] || [];
  }
}

class CarSpawner {
  constructor(direction, audioMixer, density, onCarSpawn) {
    this.direction = direction;
    this.audioMixer = audioMixer;
    this.density = density;
    this.onCarSpawn = onCarSpawn;
    this.running = false;
  }

  start() {
    this.running = true;
    const initialDelay = (this.laneIdx === 0) ? 0 : Math.random() * 4000;
    setTimeout(() => {
      if (this.running) this.spawnLoop();
    }, initialDelay);
  }

  stop() { this.running = false; }

  async spawnLoop() {
    while (this.running) {
      if (window.game && window.game.currentLaneIdx > this.laneIdx) {
        await new Promise(res => setTimeout(res, 500));
        continue;
      }
      const carFiles = AUDIO.car[this.direction];
      const fileIdx = Math.floor(Math.random() * carFiles.length);
      const file = carFiles[fileIdx];
      const vehicleMeta = createVehicleVisualMeta(fileIdx);

      const carObj = await this.audioMixer.playCar(file, this.direction, this.laneIdx, vehicleMeta);
      if (!this.running) {
        if (carObj.source) { try { carObj.source.stop(); } catch (e) { } }
        if (carObj.audio) { try { carObj.audio.pause(); } catch (e) { } }
        break;
      }
      if (this.onCarSpawn) this.onCarSpawn(carObj);
      const durationMs = (carObj.duration || 3) * 1000;
      const safeTime = durationMs * 0.90;
      const gap = Math.random() * 40;
      await new Promise(res => setTimeout(res, safeTime + gap));
    }
  }
}

// ---------- Collision Detector ---------
class CollisionDetector {
  static isCrossDanger(cars, audioCtxNow) {
    for (const car of cars) {
      let relTime = (audioCtxNow - car.startedAt);
      const duration = car.duration || 0;
      if (!duration || relTime < 0) continue;
      const dStart = car.vehicleMeta?.dangerStart ?? 0.30;
      const dEnd = car.vehicleMeta?.dangerEnd ?? 0.60;
      const dangerStart = dStart * duration;
      const dangerEnd = dEnd * duration;
      if (relTime >= dangerStart && relTime <= dangerEnd) {
        return true;
      }
    }
    return false;
  }
}

// ---------- Audio Mixer & Car Sound Engine ----------
class AudioMixer {
  constructor(laneCount, currentLane) {
    this.audio = new AudioPlayer();
    this.laneCount = laneCount;
    this.currentLane = currentLane;
    this.carObjects = [];
  }

  getVolumeForLane(laneIdx) {
    const laneDist = Math.abs(laneIdx - this.currentLane);
    if (laneDist === 0) {
      return 1;
    } else if (laneIdx < this.currentLane) {
      return 0.80;
    } else {
      if (laneDist === 1) {
        return 0.30;
      } else {
        return 0.05;
      }
    }
  }

  setCurrentLane(idx) {
    this.currentLane = idx;
    this.updateVolumes();
  }

  updateVolumes() {
    this.carObjects.forEach(car => {
      const volume = this.getVolumeForLane(car.laneIdx);
      if (car.audio && car.audio.volume !== undefined) car.audio.volume = volume;
      if (car.gainNode) car.gainNode.gain.value = volume;
    });
  }

  async playCar(file, direction, laneIdx = 0, vehicleMeta = null) {
    const volume = this.getVolumeForLane(laneIdx);
    const carObj = await this.audio.play(file, { volume });
    carObj.direction = direction;
    carObj.laneIdx = laneIdx;
    carObj.vehicleMeta = vehicleMeta || createVehicleVisualMeta(0);
    this.carObjects.push(carObj);

    if (carObj.source) carObj.source.onended = () => {
      this.carObjects = this.carObjects.filter(obj => obj !== carObj);
      (carObj.endCallback && carObj.endCallback());
    };
    if (carObj.audio) carObj.audio.addEventListener('ended', () => {
      this.carObjects = this.carObjects.filter(obj => obj !== carObj);
      (carObj.endCallback && carObj.endCallback());
    });
    return carObj;
  }

  playEvent(file, { volume = 1 } = {}) {
    return this.audio.play(file, { volume });
  }
}

// ----------- Instructions Player -----------
function playInstructions() {
  const audio = new Audio(AUDIO.instructions);
  audio.play().catch(() => {});
}

function updateInstructionsRegion(msg) {
  document.getElementById('instructions-region').innerText = msg;
}

function updateSRFeedback(msg) {
  const region = document.getElementById('screenreader-feedback');
  region.innerText = msg;
}

// ----------- Visual Renderer (3rd-Person 2.5D Cartoon Engine) -----------
class VisualRenderer {
  constructor(canvasId) {
    this.canvas = document.getElementById(canvasId);
    this.ctx = this.canvas ? this.canvas.getContext('2d') : null;
    this.width = 500;
    this.height = 420;

    // Animation States
    this.playerDisplayLane = 0;
    this.playerHopTimer = 0;
    this.bonkTimer = 0;
    this.stageBannerTimer = 0;
    this.stageBannerText = '';
    this.particles = [];
    this.soundWaves = [];
    this.demoCars = [];
    this.initDemoTraffic();
  }

  initDemoTraffic() {
    this.demoCars = [
      { lane: 0, dir: 'l', progress: 0.15, speed: 0.004, meta: createVehicleVisualMeta(0) },
      { lane: 1, dir: 'r', progress: 0.72, speed: 0.005, meta: createVehicleVisualMeta(3) },
      { lane: 2, dir: 'l', progress: 0.45, speed: 0.0032, meta: createVehicleVisualMeta(5) },
      { lane: 3, dir: 'r', progress: 0.25, speed: 0.0045, meta: createVehicleVisualMeta(1) }
    ];
  }

  triggerStepWave(stepCount) {
    this.soundWaves.push({
      radius: 10,
      maxRadius: 55 + stepCount * 15,
      alpha: 0.95,
      color: stepCount === 2 ? '#FFE600' : '#5AC8FA'
    });
  }

  triggerLaneJump(newLaneIdx) {
    this.playerHopTimer = 1.0;
    this.spawnSparkles(this.width / 2, 315, '#FFE600', 14);
  }

  triggerStageClear(stageNum) {
    this.stageBannerTimer = 1.6;
    this.stageBannerText = t('canvasStageClear')(stageNum);
    this.spawnConfetti(42);
  }

  triggerVictory() {
    this.spawnConfetti(70);
  }

  triggerBonk() {
    this.bonkTimer = 1.0;
    this.spawnSparkles(this.width / 2, 290, '#FF9500', 18);
  }

  spawnSparkles(x, y, color, count) {
    for (let i = 0; i < count; i++) {
      const angle = (Math.PI * 2 * i) / count + Math.random() * 0.4;
      const speed = 1.8 + Math.random() * 3.2;
      this.particles.push({
        type: 'star',
        x, y,
        vx: Math.cos(angle) * speed,
        vy: Math.sin(angle) * speed - 1.5,
        size: 5 + Math.random() * 4,
        color,
        life: 1.0
      });
    }
  }

  spawnConfetti(count) {
    const colors = ['#FF3B30', '#FFCC00', '#34C759', '#007AFF', '#FF2D55', '#5AC8FA'];
    for (let i = 0; i < count; i++) {
      this.particles.push({
        type: 'confetti',
        x: Math.random() * this.width,
        y: -10 - Math.random() * 80,
        vx: (Math.random() - 0.5) * 2.5,
        vy: 2 + Math.random() * 3.5,
        rot: Math.random() * Math.PI * 2,
        vRot: (Math.random() - 0.5) * 0.2,
        size: 7 + Math.random() * 5,
        color: colors[i % colors.length],
        life: 1.5
      });
    }
  }

  startLoop() {
    let lastTime = performance.now();
    const tick = (now) => {
      const dt = Math.min((now - lastTime) / 1000, 0.1);
      lastTime = now;
      this.update(dt);
      this.render(now / 1000);
      requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  }

  update(dt) {
    const game = window.game;
    const targetLane = game ? game.currentLaneIdx : 0;
    this.playerDisplayLane += (targetLane - this.playerDisplayLane) * Math.min(1, dt * 10);

    if (this.playerHopTimer > 0) this.playerHopTimer = Math.max(0, this.playerHopTimer - dt * 2.8);
    if (this.bonkTimer > 0) this.bonkTimer = Math.max(0, this.bonkTimer - dt * 0.9);
    if (this.stageBannerTimer > 0) this.stageBannerTimer = Math.max(0, this.stageBannerTimer - dt);

    // Update sound waves
    for (let i = this.soundWaves.length - 1; i >= 0; i--) {
      const w = this.soundWaves[i];
      w.radius += dt * 110;
      w.alpha -= dt * 1.4;
      if (w.alpha <= 0) this.soundWaves.splice(i, 1);
    }

    // Update particles
    for (let i = this.particles.length - 1; i >= 0; i--) {
      const p = this.particles[i];
      p.x += p.vx;
      p.y += p.vy;
      if (p.type === 'confetti') {
        p.rot += p.vRot;
        p.life -= dt * 0.45;
      } else {
        p.vy += 0.12;
        p.life -= dt * 1.5;
      }
      if (p.life <= 0) this.particles.splice(i, 1);
    }

    // Update demo traffic when on title screen
    if (!game || (!game.gameActive && !game.isPaused && !game.lossState && !game.winState)) {
      this.demoCars.forEach(dc => {
        dc.progress += dc.speed;
        if (dc.progress > 1.15) {
          dc.progress = -0.15;
          dc.meta = createVehicleVisualMeta(Math.floor(Math.random() * 6));
        }
      });
    }
  }

  // คำนวณตำแหน่งแกน Y และสเกลแบบ 2.5D Perspective จากระยะห่างของเลนเทียบกับตัวผู้เล่น
  getLaneProjection(laneOffset) {
    // laneOffset = 0 คือเลนตรงหน้าผู้เล่นพอดี (Y = 265)
    // laneOffset > 0 คือเลนที่อยู่ไกลออกไปด้านบน
    // laneOffset < 0 คือเลนที่ข้ามผ่านไปแล้วด้านล่าง
    const baseY = 268;
    const laneHeight = 46;
    const rawY = baseY - laneOffset * laneHeight;
    const depth = Math.max(-0.6, Math.min(4.5, laneOffset));
    const scale = Math.max(0.56, 1.0 - depth * 0.085);
    return { y: rawY, scale };
  }

  render(timeSec) {
    if (!this.ctx) return;
    const ctx = this.ctx;
    const game = window.game;
    const isTitle = !game || (!game.gameActive && !game.isPaused && !game.lossState && !game.winState);

    ctx.clearRect(0, 0, this.width, this.height);

    // 1. ท้องฟ้าและเมืองการ์ตูนพื้นหลัง (Sky & Thai City Skyline)
    this.drawSkyAndCity(ctx, timeSec);

    if (isTitle) {
      this.drawTitleScene(ctx, timeSec);
      return;
    }

    const totalLanes = STAGES[game.stageIdx]?.lanes || 2;
    const currentLane = game.currentLaneIdx;

    // 2. วาดพื้นถนน เลนต่างๆ ฟุตบาทต้นทาง และเป้าหมายฝั่งตรงข้าม
    this.drawRoadPerspective(ctx, totalLanes, currentLane, timeSec);

    // 3. รวบรวมรถและตัวละครมาเรียงลำดับการวาดจากไกลไปใกล้ (Painter's Algorithm)
    const drawables = [];

    // เพิ่มรถทุกคันที่กำลังวิ่งอยู่ในด่าน
    const cars = game.audioMixer?.carObjects || [];
    const audioNow = game.audioMixer?.audio?.ctx?.currentTime || (Date.now() / 1000);

    cars.forEach(car => {
      const duration = car.duration || 3;
      const elapsed = Math.max(0, audioNow - car.startedAt);
      const progress = Math.min(1.15, elapsed / duration);
      const laneOffset = car.laneIdx - this.playerDisplayLane;
      if (laneOffset >= -1.8 && laneOffset <= 5.5) {
        drawables.push({
          z: laneOffset,
          type: 'car',
          car,
          progress,
          laneOffset
        });
      }
    });

    // เพิ่มตัวละครผู้พิทักษ์ (อยู่ที่ตำแหน่งก่อนเข้าเลนปัจจุบันเล็กน้อย)
    const stepAdvance = (game.clickCounter || 0) * 0.22;
    const playerOffset = -0.35 + stepAdvance;
    drawables.push({
      z: playerOffset,
      type: 'player',
      playerOffset,
      timeSec
    });

    // เรียงจากไกล (z มาก) ไปใกล้ (z น้อย) เพื่อให้รถบังหรือถูกบังอย่างสมจริง
    drawables.sort((a, b) => b.z - a.z);

    drawables.forEach(item => {
      if (item.type === 'car') {
        this.drawCarItem(ctx, item.car, item.progress, item.laneOffset, timeSec);
      } else {
        this.drawGuardianCharacter(ctx, item.playerOffset, timeSec, game);
      }
    });

    // 4. วาดเอฟเฟกต์คลื่นเสียงและพลุ/ดาว
    this.drawEffects(ctx);

    // 5. วาดแถบสถานะบนจอ (HUD) และป้ายแจ้งเตือนผ่านด่าน/ชน/ชนะ
    this.drawHUD(ctx, game);
  }

  drawSkyAndCity(ctx, timeSec) {
    // ไล่สีท้องฟ้าสดใส
    const skyGrad = ctx.createLinearGradient(0, 0, 0, 125);
    skyGrad.addColorStop(0, '#38BDF8');
    skyGrad.addColorStop(1, '#BAE6FD');
    ctx.fillStyle = skyGrad;
    ctx.fillRect(0, 0, this.width, 125);

    // ก้อนเมฆการ์ตูนลอยช้าๆ
    ctx.fillStyle = 'rgba(255, 255, 255, 0.85)';
    const cloudX1 = ((timeSec * 12) % 600) - 80;
    const cloudX2 = ((timeSec * 8 + 260) % 600) - 80;
    this.drawCloud(ctx, cloudX1, 34, 0.9);
    this.drawCloud(ctx, cloudX2, 54, 0.7);

    // ตึกเมืองไทยสไตล์การ์ตูนที่เส้นขอบฟ้า
    const buildings = [
      { x: 10, w: 58, h: 65, color: '#FDE68A' },
      { x: 74, w: 48, h: 82, color: '#FCA5A5' },
      { x: 128, w: 66, h: 55, color: '#A7F3D0' },
      { x: 310, w: 60, h: 76, color: '#DDD6FE' },
      { x: 376, w: 54, h: 60, color: '#FDBA74' },
      { x: 436, w: 58, h: 85, color: '#93C5FD' }
    ];
    buildings.forEach(b => {
      ctx.fillStyle = b.color;
      ctx.fillRect(b.x, 122 - b.h, b.w, b.h);
      ctx.strokeStyle = '#334155';
      ctx.lineWidth = 2;
      ctx.strokeRect(b.x, 122 - b.h, b.w, b.h);
    });
  }

  drawCloud(ctx, x, y, s) {
    ctx.beginPath();
    ctx.arc(x, y, 16 * s, 0, Math.PI * 2);
    ctx.arc(x + 15 * s, y - 5 * s, 20 * s, 0, Math.PI * 2);
    ctx.arc(x + 32 * s, y, 15 * s, 0, Math.PI * 2);
    ctx.fill();
  }

  drawRoadPerspective(ctx, totalLanes, currentLane, timeSec) {
    // พื้นยางมะตอยหลัก
    ctx.fillStyle = '#334155';
    ctx.fillRect(0, 120, this.width, this.height - 120);

    // คำนวณขอบเขตฟุตบาทเริ่มต้น (ก่อนเลน 0) และฟุตบาทปลายทาง (หลังเลนสุดท้าย)
    const startSidewalkProj = this.getLaneProjection(-0.65 - this.playerDisplayLane);
    const endSidewalkProj = this.getLaneProjection((totalLanes - 0.5) - this.playerDisplayLane);

    // ฟุตบาทปลายทาง (โซนปลอดภัยสีเขียวสดใสเมื่อข้ามครบ)
    const safeTop = Math.max(115, endSidewalkProj.y - 55);
    const safeBottom = Math.min(this.height, endSidewalkProj.y + 18);
    if (safeBottom > 115) {
      ctx.fillStyle = '#22C55E';
      ctx.fillRect(0, 115, this.width, Math.max(18, safeBottom - 115));
      // ขอบฟุตบาทขาว-แดงสไตล์ถนนเมืองไทย
      this.drawThaiCurb(ctx, safeBottom - 6, 6);

      // ป้ายเชียร์ฝั่งตรงข้าม
      if (safeBottom > 142) {
        ctx.fillStyle = '#FFFFFF';
        ctx.font = 'bold 13px system-ui, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(`🏁 ${t('canvasSafeZone')}`, this.width / 2, safeBottom - 16);
      }
    }

    // วาดแต่ละเลนและเส้นแบ่งเลน
    for (let i = 0; i < totalLanes; i++) {
      const offset = i - this.playerDisplayLane;
      const proj = this.getLaneProjection(offset);
      const bandH = 46 * proj.scale;
      const topY = proj.y - bandH * 0.55;

      if (topY > this.height || topY + bandH < 118) continue;

      // ไฮไลต์เลนปัจจุบันที่กำลังจะข้ามด้วยสีเหลืองเรืองแสงนุ่มๆ (ช่วยผู้เล่นสายตาเลือนราง)
      if (i === currentLane && window.game?.gameActive) {
        const pulse = 0.16 + Math.sin(timeSec * 6) * 0.06;
        ctx.fillStyle = `rgba(255, 230, 0, ${pulse})`;
        ctx.fillRect(0, topY, this.width, bandH);
      } else if (i < currentLane) {
        // เลนที่ข้ามผ่านมาแล้ว
        ctx.fillStyle = 'rgba(34, 197, 94, 0.10)';
        ctx.fillRect(0, topY, this.width, bandH);
      }

      // เส้นทางม้าลายจางๆ ตรงกลางจอเป็นแนวสายตา
      ctx.fillStyle = 'rgba(255, 255, 255, 0.28)';
      const zebraW = 54 * proj.scale;
      ctx.fillRect((this.width - zebraW) / 2, topY + 6 * proj.scale, zebraW, bandH - 12 * proj.scale);

      // เส้นประแบ่งเลนด้านบนของเลน
      ctx.strokeStyle = i === Math.floor(totalLanes / 2) - 1 ? '#FACC15' : 'rgba(255,255,255,0.65)';
      ctx.lineWidth = Math.max(1.5, 2.5 * proj.scale);
      ctx.setLineDash([18 * proj.scale, 14 * proj.scale]);
      ctx.beginPath();
      ctx.moveTo(0, topY);
      ctx.lineTo(this.width, topY);
      ctx.stroke();
      ctx.setLineDash([]);

      // ลูกศรบอกทิศทางรถวิ่งในเลน (สลับข้างให้ตรงกับเสียง)
      const dir = i < Math.floor(totalLanes / 2) ? 'r' : 'l';
      ctx.fillStyle = 'rgba(255, 255, 255, 0.35)';
      ctx.font = `bold ${Math.round(14 * proj.scale)}px system-ui`;
      ctx.textAlign = 'center';
      const arrowChar = dir === 'l' ? '➔' : '⬅';
      ctx.fillText(arrowChar, 45, proj.y + 4);
      ctx.fillText(arrowChar, this.width - 45, proj.y + 4);
    }

    // ฟุตบาทเริ่มต้นด้านล่าง
    if (startSidewalkProj.y < this.height + 20) {
      const swY = Math.max(120, startSidewalkProj.y);
      ctx.fillStyle = '#94A3B8';
      ctx.fillRect(0, swY, this.width, this.height - swY);
      this.drawThaiCurb(ctx, swY, 7);
    }
  }

  drawThaiCurb(ctx, y, h) {
    const segW = 28;
    for (let x = 0; x < this.width; x += segW) {
      ctx.fillStyle = (Math.floor(x / segW) % 2 === 0) ? '#EF4444' : '#FFFFFF';
      ctx.fillRect(x, y, segW, h);
    }
  }

  drawCarItem(ctx, car, progress, laneOffset, timeSec) {
    const proj = this.getLaneProjection(laneOffset);
    const meta = car.vehicleMeta || VEHICLE_CATALOG[0];

    // คำนวณตำแหน่ง X ให้จุดกึ่งกลางรถผ่านกลางจอ (X = 250) ที่ช่วง progress ~ 0.45 (ตรงกับช่วงอันตราย 30%-60%)
    const travelSpan = this.width + 180;
    let x;
    if (car.direction === 'l') {
      // วิ่งจากซ้ายไปขวา (สลับฝั่งภาพให้ตรงกับเสียง)
      x = -90 + (progress / 0.90) * travelSpan;
    } else {
      // วิ่งจากขวาไปซ้าย (สลับฝั่งภาพให้ตรงกับเสียง)
      x = (this.width + 90) - (progress / 0.90) * travelSpan;
    }

    const bounce = Math.sin(timeSec * 24 + meta.wheelPhase) * 1.4 * proj.scale;
    this.drawCartoonVehicle(ctx, x, proj.y + bounce, proj.scale, car.direction, meta, timeSec);
  }

  drawCartoonVehicle(ctx, x, y, scale, dir, meta, timeSec) {
    ctx.save();
    ctx.translate(x, y);
    ctx.scale(scale * (dir === 'r' ? -1 : 1), scale);

    const w = meta.width;
    const h = meta.height;
    const c = meta.color;

    // เงาใต้ท้องรถ
    ctx.fillStyle = 'rgba(0, 0, 0, 0.35)';
    ctx.beginPath();
    ctx.ellipse(0, 6, w * 0.48, 8, 0, 0, Math.PI * 2);
    ctx.fill();

    ctx.lineWidth = 2.2;
    ctx.strokeStyle = '#1E293B';

    if (meta.type === 'moto_female' || meta.type === 'moto_male') {
      // --- วาดรถมอเตอร์ไซค์ + ผู้ขับขี่การ์ตูน ---
      const isFemale = meta.type === 'moto_female';

      // ล้อหน้าและล้อหลัง
      this.drawWheel(ctx, -14, 2, 7.5, timeSec);
      this.drawWheel(ctx, 14, 2, 7.5, timeSec);

      // ตัวถังมอเตอร์ไซค์
      ctx.fillStyle = c.body;
      ctx.beginPath();
      ctx.roundRect(-16, -12, 32, 12, 6);
      ctx.fill();
      ctx.stroke();

      // แฮนด์และไฟหน้า
      ctx.fillStyle = '#FFE600';
      ctx.beginPath();
      ctx.arc(16, -13, 3.5, 0, Math.PI * 2);
      ctx.fill();

      // ลำตัวผู้ขับขี่
      ctx.fillStyle = isFemale ? '#F472B6' : '#3B82F6';
      ctx.beginPath();
      ctx.roundRect(-7, -27, 13, 16, 5);
      ctx.fill();
      ctx.stroke();

      // ผมยาวปลิวสำหรับผู้หญิงขับ
      if (isFemale) {
        ctx.fillStyle = '#3F2E21';
        ctx.beginPath();
        ctx.ellipse(-11, -26, 8, 4.5, -0.25, 0, Math.PI * 2);
        ctx.fill();
      }

      // ศีรษะและหมวกกันน็อก
      ctx.fillStyle = isFemale ? '#FBCFE8' : '#1E293B';
      ctx.beginPath();
      ctx.arc(0, -32, 8.5, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();

      // กระบังหน้าหมวกกันน็อก
      ctx.fillStyle = '#BAE6FD';
      ctx.fillRect(2, -35, 6.5, 5);

    } else if (meta.type === 'bus') {
      // --- วาดรถโดยสารประจำทาง (รถเมล์ร้อน & ปรับอากาศ) ---
      ctx.fillStyle = c.body;
      ctx.beginPath();
      ctx.roundRect(-w / 2, -h + 4, w, h - 6, 8);
      ctx.fill();
      ctx.stroke();

      // แถบคาดสีตัดข้างรถเมล์
      ctx.fillStyle = c.accent;
      ctx.fillRect(-w / 2 + 2, -20, w - 4, 7);

      // หน้าต่างเรียงยาว
      const winCount = 5;
      const winW = (w - 26) / winCount;
      for (let i = 0; i < winCount; i++) {
        ctx.fillStyle = meta.isAirConBus ? '#38BDF8' : '#E2E8F0';
        ctx.fillRect(-w / 2 + 10 + i * (winW + 2), -h + 11, winW, 16);
        ctx.strokeRect(-w / 2 + 10 + i * (winW + 2), -h + 11, winW, 16);
      }

      // ป้ายไฟหน้ารถเมล์
      ctx.fillStyle = '#FACC15';
      ctx.fillRect(w / 2 - 14, -h + 8, 11, 7);

      // ล้อรถเมล์
      this.drawWheel(ctx, -w * 0.30, 2, 9.5, timeSec);
      this.drawWheel(ctx, w * 0.30, 2, 9.5, timeSec);

    } else if (meta.type === 'van') {
      // --- วาดรถตู้ 10-12 ที่นั่ง ---
      ctx.fillStyle = c.body;
      ctx.beginPath();
      ctx.roundRect(-w / 2, -h + 4, w, h - 6, 10);
      ctx.fill();
      ctx.stroke();

      // แถบคาดข้างรถตู้
      ctx.fillStyle = c.accent;
      ctx.fillRect(-w / 2 + 2, -18, w - 4, 5);

      // หน้าต่าง 3 บานใหญ่
      for (let i = 0; i < 3; i++) {
        ctx.fillStyle = '#BAE6FD';
        ctx.fillRect(-w / 2 + 12 + i * 26, -h + 10, 22, 15);
        ctx.strokeRect(-w / 2 + 12 + i * 26, -h + 10, 22, 15);
      }

      this.drawWheel(ctx, -w * 0.28, 2, 8.5, timeSec);
      this.drawWheel(ctx, w * 0.28, 2, 8.5, timeSec);

    } else if (meta.type === 'suv7') {
      // --- วาดรถยนต์ครอบครัว 7 ที่นั่ง (SUV) ---
      ctx.fillStyle = c.body;
      ctx.beginPath();
      ctx.roundRect(-w / 2, -24, w, 22, 7);
      ctx.fill();
      ctx.stroke();

      // ห้องโดยสารทรงสูงยาว + แร็คหลังคา
      ctx.fillStyle = c.body;
      ctx.beginPath();
      ctx.roundRect(-w * 0.38, -h + 4, w * 0.72, 20, 6);
      ctx.fill();
      ctx.stroke();

      // กระจก 3 ช่อง
      ctx.fillStyle = '#BAE6FD';
      ctx.fillRect(-w * 0.33, -h + 8, 17, 12);
      ctx.fillRect(-w * 0.10, -h + 8, 19, 12);
      ctx.fillRect(w * 0.14, -h + 8, 15, 12);

      this.drawWheel(ctx, -w * 0.28, 2, 8.5, timeSec);
      this.drawWheel(ctx, w * 0.28, 2, 8.5, timeSec);

    } else {
      // --- วาดรถยนต์เก๋ง 4 ที่นั่ง (Sedan) ---
      ctx.fillStyle = c.body;
      ctx.beginPath();
      ctx.roundRect(-w / 2, -20, w, 18, 8);
      ctx.fill();
      ctx.stroke();

      // หลังคาโค้งมน
      ctx.fillStyle = c.body;
      ctx.beginPath();
      ctx.roundRect(-w * 0.28, -h + 3, w * 0.56, 16, 8);
      ctx.fill();
      ctx.stroke();

      // กระจกหน้า-หลัง
      ctx.fillStyle = '#BAE6FD';
      ctx.fillRect(-w * 0.22, -h + 6, 16, 10);
      ctx.fillRect(2, -h + 6, 16, 10);

      this.drawWheel(ctx, -w * 0.28, 2, 7.8, timeSec);
      this.drawWheel(ctx, w * 0.28, 2, 7.8, timeSec);
    }

    // ไฟหน้าสว่าง (สำหรับรถยนต์ทุกประเภท)
    if (meta.type !== 'moto_female' && meta.type !== 'moto_male') {
      ctx.fillStyle = '#FEF08A';
      ctx.beginPath();
      ctx.arc(w / 2 - 3, -11, 3.8, 0, Math.PI * 2);
      ctx.fill();
    }

    ctx.restore();
  }

  drawWheel(ctx, wx, wy, r, timeSec) {
    ctx.save();
    ctx.translate(wx, wy);
    ctx.fillStyle = '#1E293B';
    ctx.beginPath();
    ctx.arc(0, 0, r, 0, Math.PI * 2);
    ctx.fill();

    // ดุมล้อหมุน
    ctx.rotate(timeSec * 14);
    ctx.fillStyle = '#E2E8F0';
    ctx.beginPath();
    ctx.arc(0, 0, r * 0.45, 0, Math.PI * 2);
    ctx.fill();
    ctx.strokeStyle = '#64748B';
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(-r * 0.4, 0);
    ctx.lineTo(r * 0.4, 0);
    ctx.stroke();
    ctx.restore();
  }

  drawGuardianCharacter(ctx, playerOffset, timeSec, game) {
    const proj = this.getLaneProjection(playerOffset);
    const x = this.width / 2;
    const hopOffset = Math.sin(this.playerHopTimer * Math.PI) * 26;
    const bonkBounce = this.bonkTimer > 0 ? Math.sin(this.bonkTimer * Math.PI * 3) * 14 : 0;
    const idleBob = Math.sin(timeSec * 5) * 2;
    const y = proj.y - hopOffset - Math.abs(bonkBounce) + idleBob;

    ctx.save();
    ctx.translate(x, y);
    ctx.scale(proj.scale * 1.12, proj.scale * 1.12);

    // เงาใต้เท้าตัวละคร
    ctx.fillStyle = 'rgba(0, 0, 0, 0.35)';
    ctx.beginPath();
    ctx.ellipse(0, 8 + hopOffset, 18, 6, 0, 0, Math.PI * 2);
    ctx.fill();

    // ขาและรองเท้า (ก้าวสลับเวลาผู้เล่นกดข้าม)
    const stepShift = (game?.clickCounter || 0) % 2 === 1 ? 4 : -4;
    ctx.fillStyle = '#1E3A8A';
    ctx.fillRect(-10, -14 + (stepShift > 0 ? -3 : 0), 8, 18);
    ctx.fillRect(2, -14 + (stepShift < 0 ? -3 : 0), 8, 18);

    // รองเท้าผ้าใบ
    ctx.fillStyle = '#FFFFFF';
    ctx.beginPath();
    ctx.roundRect(-11, 2, 10, 6, 3);
    ctx.roundRect(1, 2, 10, 6, 3);
    ctx.fill();

    // เสื้ออาสาสมัครผู้พิทักษ์สีเหลืองสดสะท้อนแสง (มองเห็นชัดจากด้านหลัง)
    ctx.fillStyle = '#FACC15';
    ctx.strokeStyle = '#1E293B';
    ctx.lineWidth = 2.2;
    ctx.beginPath();
    ctx.roundRect(-14, -40, 28, 28, 8);
    ctx.fill();
    ctx.stroke();

    // กระเป๋าเป้สะพายหลังสีส้มน่ารัก
    ctx.fillStyle = '#F97316';
    ctx.beginPath();
    ctx.roundRect(-10, -37, 20, 20, 6);
    ctx.fill();
    ctx.stroke();

    // แถบสะท้อนแสงบนกระเป๋าเป้
    ctx.fillStyle = '#FEF08A';
    ctx.fillRect(-7, -30, 14, 4);

    // ศีรษะหันหลัง (ทรงผมสีน้ำตาลเข้ม + หูฟัง/หูตั้งใจฟังเสียง)
    ctx.fillStyle = '#FDBA74';
    ctx.beginPath();
    ctx.arc(0, -50, 12, 0, Math.PI * 2);
    ctx.fill();
    ctx.stroke();

    // เส้นผมด้านหลังศีรษะ
    ctx.fillStyle = '#292524';
    ctx.beginPath();
    ctx.arc(0, -51, 12.2, Math.PI * 0.85, Math.PI * 2.15);
    ctx.fill();

    // ไม้เท้าขาว (White Cane) ในมือขวาชี้ไปข้างหน้า
    ctx.strokeStyle = '#FFFFFF';
    ctx.lineWidth = 3.5;
    ctx.beginPath();
    ctx.moveTo(14, -22);
    ctx.lineTo(28, -42);
    ctx.stroke();
    // ปลายไม้เท้าสีแดง
    ctx.strokeStyle = '#EF4444';
    ctx.beginPath();
    ctx.moveTo(24, -36);
    ctx.lineTo(28, -42);
    ctx.stroke();

    // หากกำลังตั้งใจฟัง (กดไปแล้ว 1 หรือ 2 ครั้ง) วาดสัญลักษณ์ตั้งใจฟังเสียง
    if (game?.gameActive && game.clickCounter > 0 && game.clickCounter < 3) {
      ctx.fillStyle = '#FFE600';
      ctx.font = 'bold 14px system-ui';
      ctx.textAlign = 'center';
      ctx.fillText(`👂 ${game.clickCounter}/3`, 0, -70);
    }

    // หากโดนชน วาดดาวและนกบินวนเหนือหัวแบบการ์ตูนน่ารัก
    if (game?.lossState || this.bonkTimer > 0) {
      for (let i = 0; i < 3; i++) {
        const ang = timeSec * 5 + (i * Math.PI * 2) / 3;
        const sx = Math.cos(ang) * 20;
        const sy = -68 + Math.sin(ang) * 6;
        ctx.fillStyle = '#FACC15';
        ctx.font = '15px system-ui';
        ctx.fillText('⭐', sx - 7, sy);
      }
    }

    ctx.restore();
  }

  drawEffects(ctx) {
    // วงคลื่นเสียงรอบตัวละครตอนกดก้าวขา
    const centerX = this.width / 2;
    const centerY = 275;
    this.soundWaves.forEach(w => {
      ctx.save();
      ctx.strokeStyle = w.color;
      ctx.globalAlpha = Math.max(0, w.alpha);
      ctx.lineWidth = 3;
      ctx.beginPath();
      ctx.arc(centerX, centerY, w.radius, 0, Math.PI * 2);
      ctx.stroke();
      ctx.restore();
    });

    // ดาวและพลุกระดาษ
    this.particles.forEach(p => {
      ctx.save();
      ctx.globalAlpha = Math.min(1, p.life);
      ctx.translate(p.x, p.y);
      if (p.type === 'confetti') {
        ctx.rotate(p.rot);
        ctx.fillStyle = p.color;
        ctx.fillRect(-p.size / 2, -p.size / 3, p.size, p.size * 0.6);
      } else {
        ctx.fillStyle = p.color;
        ctx.beginPath();
        ctx.arc(0, 0, p.size * 0.5, 0, Math.PI * 2);
        ctx.fill();
      }
      ctx.restore();
    });
  }

  drawHUD(ctx, game) {
    const totalStages = STAGES.length;
    const stageNum = Math.min(game.stageIdx + 1, totalStages);
    const totalLanes = STAGES[Math.min(game.stageIdx, totalStages - 1)].lanes;
    const laneNum = Math.min(game.currentLaneIdx + 1, totalLanes);

    let crossed = 0;
    for (let i = 0; i < game.stageIdx; i++) crossed += STAGES[i].lanes;
    if (!game.winState) crossed += game.currentLaneIdx;
    const score = crossed * 25;

    // แถบ HUD ด้านบนซ้าย-ขวา
    ctx.fillStyle = 'rgba(15, 23, 42, 0.78)';
    ctx.beginPath();
    ctx.roundRect(12, 12, 195, 34, 10);
    ctx.roundRect(this.width - 142, 12, 130, 34, 10);
    ctx.fill();

    ctx.fillStyle = '#FFE600';
    ctx.font = 'bold 14px system-ui, sans-serif';
    ctx.textAlign = 'left';
    ctx.fillText(`${t('canvasStageHud')(stageNum, totalStages)} • ${t('canvasLaneHud')(laneNum, totalLanes)}`, 22, 34);

    ctx.fillStyle = '#4ADE80';
    ctx.textAlign = 'right';
    ctx.fillText(t('canvasScoreHud')(score), this.width - 22, 34);

    // ป้ายแบนเนอร์ผ่านด่าน / ชน / ชนะกลางจอ
    if (this.stageBannerTimer > 0) {
      this.drawCenterBanner(ctx, this.stageBannerText, '#16A34A', '#FFFFFF');
    } else if (game.lossState) {
      this.drawCenterBanner(ctx, t('canvasBonk'), '#DC2626', '#FFE600');
    } else if (game.winState) {
      this.drawCenterBanner(ctx, `🏆 ${t('canvasVictory')}`, '#D97706', '#FFFFFF');
    }
  }

  drawCenterBanner(ctx, text, bgColor, textColor) {
    ctx.save();
    ctx.fillStyle = bgColor;
    ctx.strokeStyle = '#FFFFFF';
    ctx.lineWidth = 3;
    ctx.beginPath();
    ctx.roundRect(55, 175, this.width - 110, 56, 16);
    ctx.fill();
    ctx.stroke();

    ctx.fillStyle = textColor;
    ctx.font = 'bold 21px system-ui, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(text, this.width / 2, 210);
    ctx.restore();
  }

  drawTitleScene(ctx, timeSec) {
    // ถนนหน้าแรกพร้อมรถหลากสีวิ่งผ่านเบาๆ
    ctx.fillStyle = '#334155';
    ctx.fillRect(0, 122, this.width, 298);

    // เส้นประถนนหน้าแรก
    for (let row = 0; row < 4; row++) {
      const y = 165 + row * 48;
      ctx.strokeStyle = row === 1 ? '#FACC15' : 'rgba(255,255,255,0.6)';
      ctx.lineWidth = 2.5;
      ctx.setLineDash([18, 14]);
      ctx.beginPath();
      ctx.moveTo(0, y);
      ctx.lineTo(this.width, y);
      ctx.stroke();
      ctx.setLineDash([]);
    }

    // รถ Demo วิ่งผ่านหน้าแรก (สลับข้างให้สอดคล้องกับในเกม)
    this.demoCars.forEach(dc => {
      const y = 148 + dc.lane * 48;
      const x = dc.dir === 'l'
        ? -80 + dc.progress * (this.width + 160)
        : (this.width + 80) - dc.progress * (this.width + 160);
      this.drawCartoonVehicle(ctx, x, y, 0.88, dc.dir, dc.meta, timeSec);
    });

    // ฟุตบาทหน้าแรก + ตัวละครยืนรอ
    ctx.fillStyle = '#94A3B8';
    ctx.fillRect(0, 325, this.width, 95);
    this.drawThaiCurb(ctx, 325, 7);

    // วงคลื่นเสียงเบาๆ รอบตัวละคร
    const pulseR = 16 + ((timeSec * 24) % 36);
    const pulseAlpha = 1 - ((timeSec * 24) % 36) / 36;
    ctx.strokeStyle = `rgba(255, 230, 0, ${pulseAlpha * 0.8})`;
    ctx.lineWidth = 2.5;
    ctx.beginPath();
    ctx.arc(this.width / 2, 345, pulseR, 0, Math.PI * 2);
    ctx.stroke();

    this.drawGuardianCharacter(ctx, -1.65, timeSec, null);

    // ป้ายชื่อเกมบนจอภาพ
    ctx.fillStyle = 'rgba(15, 23, 42, 0.82)';
    ctx.beginPath();
    ctx.roundRect(45, 24, this.width - 90, 74, 16);
    ctx.fill();

    ctx.fillStyle = '#FFE600';
    ctx.font = 'bold 22px system-ui, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(t('canvasTitle'), this.width / 2, 56);

    ctx.fillStyle = '#F8FAFC';
    ctx.font = '14px system-ui, sans-serif';
    ctx.fillText(t('canvasSubtitle'), this.width / 2, 82);
  }
}

// ----------------- UI Handler & Game Logic -----------------
class CrossroadGame {
  dangerInterval = null;

  startDangerCheck() {
    if (this.dangerInterval) return;
    this.dangerInterval = setInterval(() => {
      if (!this.gameActive) return;
      if (this.clickCounter > 0 && this.clickCounter < CONTROLS_REQUIRED_CLICKS) {
        const now = this.audioMixer.audio.ctx?.currentTime || (Date.now() / 1000);
        const carsInLane = this.laneManager.getActiveCarsInLane(this.currentLaneIdx);
        if (CollisionDetector.isCrossDanger(carsInLane, now)) {
          clearInterval(this.dangerInterval);
          this.dangerInterval = null;
          this.endGame(true);
        }
      }
    }, 60);
  }

  stopDangerCheck() {
    if (this.dangerInterval) {
      clearInterval(this.dangerInterval);
      this.dangerInterval = null;
    }
  }

  constructor() {
    this.stageIdx = 0;
    this.currentLaneIdx = 0;
    this.clickCounter = 0;
    this.laneManager = null;
    this.audioMixer = null;
    this.gameActive = false;
    this.isPaused = false;
    this.lossState = false;
    this.winState = false;
    this.renderer = new VisualRenderer('game-canvas');
    this.renderer.startLoop();
  }

  setPlayingChromeVisibility(isPlaying) {
    document.body.classList.toggle('is-playing', isPlaying);

    const topBar = document.getElementById('top-bar');
    if (topBar) topBar.style.display = isPlaying ? 'none' : '';

    const storyEl = document.getElementById('story');
    if (storyEl) storyEl.style.display = isPlaying ? 'none' : '';
  }

  pauseGame() {
    if (!this.gameActive || this.isPaused) return;
    this.gameActive = false;
    this.isPaused = true;

    this.clickCounter = 0;
    this.stopDangerCheck();

    this.laneManager?.stopSpawners();
    this.audioMixer?.carObjects.forEach(car => {
      if (car.source) { try { car.source.stop(); } catch (e) { } }
      if (car.audio) { try { car.audio.pause(); } catch (e) { } }
    });
    if (this.audioMixer) this.audioMixer.carObjects = [];

    const crossBtn = document.getElementById('btn-cross');
    if (crossBtn) {
      crossBtn.removeAttribute('role');
      crossBtn.style.display = 'none';
    }

    // คืนค่าแถบด้านบน ปุ่มสลับภาษา และเรื่องราวเมื่อหยุดชั่วคราว
    this.setPlayingChromeVisibility(false);

    const startBtn = document.getElementById('btn-start-game');
    if (startBtn) {
      startBtn.innerText = t('resumeBtn');
      startBtn.setAttribute('aria-label', t('resumeBtn'));
      startBtn.style.display = '';
      startBtn.focus();
    }

    updateSRFeedback(t('pausedSR'));
  }

  resumeGame() {
    if (!this.isPaused) return;
    this.isPaused = false;
    this.gameActive = true;
    this.clickCounter = 0;

    // ซ่อนแถบด้านบนและรายละเอียดหน้าแรกกลับเข้าโหมดเล่นเกม
    this.setPlayingChromeVisibility(true);

    const startBtn = document.getElementById('btn-start-game');
    if (startBtn) startBtn.style.display = 'none';

    const crossBtn = document.getElementById('btn-cross');
    if (crossBtn) {
      crossBtn.style.display = '';
      crossBtn.setAttribute('role', 'application');
      crossBtn.focus();
    }

    this.laneManager?.stopSpawners();
    this.audioMixer = new AudioMixer(STAGES[this.stageIdx].lanes, this.currentLaneIdx);
    this.laneManager = new LaneManager(this.stageIdx);
    this.laneManager.startSpawners(this.audioMixer);
    this.audioMixer.setCurrentLane(this.currentLaneIdx);

    this.renderStatus();
  }

  returnToMainMenu() {
    if (!this.isPaused && !this.lossState && !this.winState) return;

    this.gameActive = false;
    this.isPaused = false;
    this.lossState = false;
    this.winState = false;
    this.clickCounter = 0;
    this.stageIdx = 0;
    this.currentLaneIdx = 0;
    this.stopDangerCheck();

    this.laneManager?.stopSpawners();
    this.audioMixer?.carObjects.forEach(car => {
      if (car.source) { try { car.source.stop(); } catch (e) { } }
      if (car.audio) { try { car.audio.pause(); } catch (e) { } }
    });
    if (this.audioMixer) this.audioMixer.carObjects = [];

    if (this.renderer) {
      this.renderer.playerDisplayLane = 0;
      this.renderer.bonkTimer = 0;
      this.renderer.stageBannerTimer = 0;
    }

    this.hideRestart();

    const resultDiv = document.getElementById('game-result');
    if (resultDiv) resultDiv.style.display = 'none';

    const crossBtn = document.getElementById('btn-cross');
    if (crossBtn) {
      crossBtn.removeAttribute('role');
      crossBtn.style.display = 'none';
    }

    updateInstructionsRegion('');
    this.setPlayingChromeVisibility(false);

    const startBtn = document.getElementById('btn-start-game');
    if (startBtn) {
      startBtn.innerText = t('startBtn');
      startBtn.setAttribute('aria-label', t('startBtn'));
      startBtn.style.display = '';
      startBtn.focus();
    }

    updateSRFeedback(t('returnedMainSR'));
  }

  async start() {
    // ซ่อนปุ่มย้อนกลับ ปุ่มสลับภาษา และส่วนเรื่องราว/วิธีเล่นระหว่างเล่นเกม
    this.setPlayingChromeVisibility(true);

    const resultDiv = document.getElementById('game-result');
    if (resultDiv) resultDiv.style.display = 'none';

    const startBtn = document.getElementById('btn-start-game');
    if (startBtn) {
      startBtn.innerText = t('startBtn');
      startBtn.setAttribute('aria-label', t('startBtn'));
      startBtn.style.display = 'none';
    }

    const crossBtn = document.getElementById('btn-cross');
    if (crossBtn) {
      crossBtn.style.display = '';
      crossBtn.setAttribute('role', 'application');
      crossBtn.focus();
    }

    updateInstructionsRegion(t('instIntro'));
    playInstructions();
    this.stageIdx = 0;
    this.currentLaneIdx = 0;
    this.clickCounter = 0;
    this.gameActive = true;
    this.isPaused = false;
    this.lossState = false;
    this.winState = false;
    this.renderer.playerDisplayLane = 0;
    this.hideRestart();
    this.audioMixer = new AudioMixer(STAGES[this.stageIdx].lanes, this.currentLaneIdx);
    this.laneManager = new LaneManager(this.stageIdx);
    this.laneManager.startSpawners(this.audioMixer);
    this.renderStatus();
  }

  updateEndResultUI(announceSR = true) {
    let totalLanes = 0;
    for (let i = 0; i < STAGES.length; i++) {
      totalLanes += STAGES[i].lanes;
    }
    let crossed = 0;
    if (this.lossState) {
      for (let i = 0; i < this.stageIdx; i++) {
        crossed += STAGES[i].lanes;
      }
      crossed += this.currentLaneIdx;
    } else {
      crossed = totalLanes;
    }
    let score = crossed * 25;
    let roadsCrossed = this.lossState ? this.stageIdx : STAGES.length;
    let summary = t('summary')(roadsCrossed, crossed, totalLanes, score);

    if (announceSR) {
      if (this.lossState) {
        updateSRFeedback(t('srLoss')(summary));
      } else {
        updateSRFeedback(t('srWin')(summary));
      }
    }

    const resultDiv = document.getElementById('game-result');
    if (resultDiv) {
      resultDiv.innerText = summary;
      resultDiv.style.display = '';
    }
  }

  endGame(loss) {
    this.gameActive = false;
    this.isPaused = false;
    this.lossState = loss;
    this.winState = !loss;
    this.laneManager?.stopSpawners();
    this.audioMixer?.carObjects.forEach(car => {
      if (car.source) { try { car.source.stop(); } catch (e) { } }
      if (car.audio) { try { car.audio.pause(); } catch (e) { } }
    });
    this.audioMixer?.playEvent(loss ? AUDIO.crash : AUDIO.win);

    if (loss) {
      this.renderer?.triggerBonk();
    } else {
      this.renderer?.triggerVictory();
    }

    this.showRestart();

    // คืนค่าแถบด้านบน ปุ่มสลับภาษา และเรื่องราว/วิธีเล่นกลับมาในหน้าจบเกม
    this.setPlayingChromeVisibility(false);

    this.updateEndResultUI(true);
  }

  showRestart() { document.getElementById('btn-restart').style.display = ''; }
  hideRestart() { document.getElementById('btn-restart').style.display = 'none'; }

  async handleCross() {
    if (!this.gameActive) return;

    this.clickCounter++;
    if (this.clickCounter < CONTROLS_REQUIRED_CLICKS) {
      updateSRFeedback(t('pressMore')(CONTROLS_REQUIRED_CLICKS - this.clickCounter));
      this.renderer?.triggerStepWave(this.clickCounter);
      if (this.clickCounter === 1) {
        this.startDangerCheck();
      }
      return;
    }

    this.stopDangerCheck();
    this.clickCounter = 0;

    const now = this.audioMixer.audio.ctx?.currentTime || (Date.now() / 1000);
    const carsInLane = this.laneManager.getActiveCarsInLane(this.currentLaneIdx);

    if (CollisionDetector.isCrossDanger(carsInLane, now)) {
      this.endGame(true);
      return;
    }

    // ข้ามเลนสำเร็จ!
    this.audioMixer.playEvent(AUDIO.lanePass);
    let crossed = 0;
    for (let i = 0; i < this.stageIdx; i++) {
      crossed += STAGES[i].lanes;
    }
    crossed += Math.min(this.currentLaneIdx + 1, STAGES[this.stageIdx].lanes);
    let totalLanes = 0;
    for (let i = 0; i < STAGES.length; i++) {
      totalLanes += STAGES[i].lanes;
    }
    let score = crossed * 25;
    let roadsCrossed = this.stageIdx;
    updateSRFeedback(t('lanePassed')(this.currentLaneIdx + 1, roadsCrossed, crossed, score));

    this.laneManager.stopSpawnerForLane(this.currentLaneIdx);
    this.currentLaneIdx++;
    this.renderer?.triggerLaneJump(this.currentLaneIdx);

    if (this.currentLaneIdx >= STAGES[this.stageIdx].lanes) {
      // ผ่านด่าน!
      const clearedStageNumber = this.stageIdx + 1;
      this.stageIdx++;
      if (this.stageIdx >= STAGES.length) {
        this.endGame(false);
        return;
      }
      this.renderer?.triggerStageClear(clearedStageNumber);
      updateSRFeedback(t('stageReach')(this.stageIdx, STAGES[this.stageIdx].lanes));
      this.currentLaneIdx = 0;
      this.renderer.playerDisplayLane = 0;
      this.laneManager.stopSpawners();
      this.audioMixer = new AudioMixer(STAGES[this.stageIdx].lanes, this.currentLaneIdx);
      this.laneManager = new LaneManager(this.stageIdx);
      this.laneManager.startSpawners(this.audioMixer);
      this.startDangerCheck();
      setTimeout(() => updateSRFeedback(t('stageReady')), 1300);
    }
    this.audioMixer.setCurrentLane(this.currentLaneIdx);
    this.renderStatus();
  }

  renderStatus() {
    updateSRFeedback(this.winState
      ? t('statusWin')
      : (this.lossState
        ? t('statusLoss')
        : t('statusPlaying')(this.stageIdx + 1, STAGES.length, this.currentLaneIdx + 1, STAGES[this.stageIdx].lanes))
    );
    document.getElementById('btn-cross').setAttribute('aria-label',
      t('crossAriaPlaying')(this.currentLaneIdx + 1, STAGES[this.stageIdx].lanes));
  }

  initControls() {
    const pressedKeys = new Set();
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        if (this.gameActive) {
          e.preventDefault();
          this.pauseGame();
          return;
        }
        if (this.isPaused || this.lossState || this.winState) {
          e.preventDefault();
          this.returnToMainMenu();
          return;
        }
      }
      if ((e.key === 'r' || e.key === 'R') && !this.gameActive && (this.isPaused || this.lossState || this.winState)) {
        updateSRFeedback(t('restarting'));
        this.start();
        return;
      }
      if (!this.gameActive) return;
      if (e.key === 'ArrowUp') {
        if (pressedKeys.has('ArrowUp')) return;
        pressedKeys.add('ArrowUp');
        this.handleCross();
      }
    });
    window.addEventListener('keyup', (e) => {
      if (e.key === 'ArrowUp') {
        pressedKeys.delete('ArrowUp');
      }
    });

    const crossBtn = document.getElementById('btn-cross');
    if (crossBtn) {
      crossBtn.addEventListener('click', () => {
        if (this.gameActive) {
          this.handleCross();
        }
      });

      crossBtn.addEventListener('touchend', (e) => {
        e.preventDefault();
        this.handleCross();
      });
    }

    document.getElementById('btn-restart').addEventListener('click', () => {
      updateSRFeedback(t('restarting'));
      this.start();
    });

    Array.from(document.querySelectorAll('button')).forEach(btn => {
      btn.addEventListener('focus', () => {
        btn.setAttribute('aria-live', 'polite');
      });
      btn.addEventListener('blur', () => {
        btn.removeAttribute('aria-live');
      });
    });
  }
}

// -- Instantiate the game --
window.addEventListener('DOMContentLoaded', () => {
  applyLanguageUI();
  window.game = new CrossroadGame();
  window.game.initControls();
  const startBtn = document.getElementById('btn-start-game');
  startBtn.addEventListener('click', () => {
    if (window.game.isPaused) {
      window.game.resumeGame();
    } else {
      window.game.start();
    }
  });
});