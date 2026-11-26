// === Crossroad: Audio-Only Road-Crossing Game ===
/*
  Architecture:
    - LaneManager: Handles lanes, car spawning, and per-lane logic
    - CarSpawner: Controls timing and spawning of cars in lanes
    - CollisionDetector: Checks for danger zone collisions
    - AudioMixer: Handles directional mixing, volume, panning, overlapping
    - UIHandler: Controls, instructions, ARIA feedback, restart, etc.
    - InstructionsPlayer: Controls initial voice instructions
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
    pausedSR: 'หยุดเกมชั่วคราว ออกจากโหมดโฟกัสแล้ว กดปุ่ม เล่นต่อ เพื่อดำเนินการต่อ',
    pressMore: (n) => `กดอีก ${n} ครั้งเพื่อข้ามเลน`,
    lanePassed: (lane, roads, crossed, score) => `ข้ามเลนที่ ${lane} สำเร็จ! คุณข้ามถนนได้ ${roads} เส้น รวม ${crossed} เลน คะแนนรวม ${score} คะแนน`,
    stageReach: (stage, lanes) => `ถึงด่านที่ ${stage} แล้ว ตอนนี้มี ${lanes} เลน`,
    stageReady: 'เตรียมตัวสำหรับด่านถัดไป!',
    statusWin: 'เกมจบแล้ว! คุณชนะ',
    statusLoss: 'คุณถูกรถชน ลองใหม่หรือฟังคำแนะนำอีกครั้ง',
    statusPlaying: (s, totalS, l, totalL) => `ด่านที่ ${s}/${totalS}, เลนที่ ${l}/${totalL}`,
    crossAriaPlaying: (l, totalL) => `ข้ามเลน: กด 3 ครั้งเพื่อข้ามเลนที่ ${l} จาก ${totalL}`,
    summary: (roads, crossed, totalL, score) => `คุณข้ามถนนได้ ${roads} เส้น รวม ${crossed} เลน (จากทั้งหมด ${totalL} เลน) คะแนนรวม ${score} คะแนน`,
    srLoss: (sum) => `ชน! คุณถูกรถชน ${sum} กด R หรือปุ่มเริ่มใหม่เพื่อเล่นอีกครั้ง`,
    srWin: (sum) => `ยินดีด้วย! คุณข้ามถนนได้สำเร็จและชนะแล้ว ${sum}`,
    restarting: 'กำลังเริ่มเกมใหม่...',
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
    pausedSR: 'Game paused. Exited focus mode. Press Resume to continue.',
    pressMore: (n) => `Press ${n} more time${n > 1 ? 's' : ''} to cross the lane`,
    lanePassed: (lane, roads, crossed, score) => `Lane ${lane} crossed! You crossed ${roads} road(s), ${crossed} lane(s) total. Total score: ${score} points`,
    stageReach: (stage, lanes) => `Reached Stage ${stage}. Now there are ${lanes} lanes.`,
    stageReady: 'Get ready for the next stage!',
    statusWin: 'Game Over! You Win!',
    statusLoss: 'You were hit by a car. Try again or read the instructions again.',
    statusPlaying: (s, totalS, l, totalL) => `Stage ${s}/${totalS}, Lane ${l}/${totalL}`,
    crossAriaPlaying: (l, totalL) => `Cross lane: Press 3 times to cross lane ${l} of ${totalL}`,
    summary: (roads, crossed, totalL, score) => `You crossed ${roads} road(s), ${crossed} lane(s) (out of ${totalL} total lanes). Total score: ${score} points`,
    srLoss: (sum) => `Crash! You were hit by a car. ${sum}. Press R or the Restart button to play again.`,
    srWin: (sum) => `Congratulations! You crossed the roads safely and won! ${sum}`,
    restarting: 'Restarting game...',
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

// -------- Audio Asset Paths --------
const AUDIO = {
  car: {
    l: ['audio/car-l1.mp3', 'audio/car-l2.mp3', 'audio/car-l3.mp3', 'audio/car-l4.mp3', 'audio/car-l5.mp3', 'audio/car-l6.mp3'],
    r: ['audio/car-r1.mp3', 'audio/car-r2.mp3', 'audio/car-r3.mp3', 'audio/car-r4.mp3', 'audio/car-r5.mp3', 'audio/car-r6.mp3']
  },
  crash: 'audio/car-crash.mp3',
  lanePass: 'audio/lane-pass.mp3',
  win: 'audio/win-sound.mp3',
  instructions: 'audio/instructions.mp3' // This should be developer-supplied. Fallback: text-to-speech.
};

// ----------- Game Config -----------
const STAGES = [
  { lanes: 2, carDensity: 0.7 },   // ด่าน 1: มี 2 เลน ความหนาแน่นรถ 0.7
  { lanes: 4, carDensity: 1.0 },   // ด่าน 2: มี 4 เลน ความหนาแน่นรถ 1.0
  { lanes: 6, carDensity: 1.3 },   // ด่าน 3: มี 6 เลน ความหนาแน่นรถ 1.3
  { lanes: 8, carDensity: 1.6 },   // ด่าน 4: มี 8 เลน ความหนาแน่นรถ 1.6
  { lanes: 10, carDensity: 1.9 },  // ด่าน 5: มี 10 เลน ความหนาแน่นรถ 1.9
  { lanes: 12, carDensity: 2.2 },  // ด่าน 6: มี 12 เลน ความหนาแน่นรถ 2.2
  { lanes: 14, carDensity: 2.5 },  // ด่าน 7: มี 14 เลน ความหนาแน่นรถ 2.5
  { lanes: 16, carDensity: 2.8 }   // ด่าน 8: มี 16 เลน ความหนาแน่นรถ 2.8
];

const CONTROLS_REQUIRED_CLICKS = 3;

// ----------- Audio Helper -----------
class AudioPlayer {
  constructor() {
    this.ctx = window.AudioContext ? new AudioContext() : null;
    this.buffers = {}; // Cache for loaded audio
  }

  // Load and decode .mp3 file
  async loadBuffer(url) {
    if (this.buffers[url]) return this.buffers[url];
    const res = await fetch(url);
    const arrBuf = await res.arrayBuffer();
    const buffer = await this.ctx.decodeAudioData(arrBuf);
    this.buffers[url] = buffer;
    return buffer;
  }

  // Play an audio file with options: {volume, onEnded}
  async play(url, { volume = 1, onEnded = null } = {}) {
    if (!this.ctx) {
      const audio = new Audio(url);
      audio.volume = volume;
      audio.addEventListener('ended', () => {
        if (onEnded) onEnded();
      });
      audio.play();
      return { audio, startedAt: Date.now(), duration: audio.duration || 0, url };
    }
    const buffer = await this.loadBuffer(url);
    const source = this.ctx.createBufferSource();
    source.buffer = buffer;

    let gain = this.ctx.createGain();
    gain.gain.value = volume;

    source.connect(gain).connect(this.ctx.destination);

    // Danger tracking: stores car start time
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
      // ครึ่งแรกเล่น car-r(n), ครึ่งหลังเล่น car-l(n)
      let dir = i < halfLane ? 'r' : 'l';
      this.laneConfigs.push({ direction: dir });
    }
  }

  startSpawners(audioMixer) {
    // Car spawners per lane
    this.activeCarsPerLane = [];
    this.spawners = [];
    const density = STAGES[this.stageIdx].carDensity;
    for (let i = 0; i < this.laneConfigs.length; ++i) {
      let spawner = new CarSpawner(this.laneConfigs[i].direction, audioMixer, density, (carObj) => {
        if (!this.activeCarsPerLane[i]) this.activeCarsPerLane[i] = [];
        this.activeCarsPerLane[i].push(carObj);
        // Remove car when audio ends
        carObj.endCallback = () => {
          this.activeCarsPerLane[i] = this.activeCarsPerLane[i].filter(obj => obj !== carObj);
        };
      });
      spawner.laneIdx = i; // เพิ่ม laneIdx ให้ spawner
      spawner.start();
      this.spawners.push(spawner);
      this.activeCarsPerLane.push([]);
    }
  }

  // หยุด CarSpawner เฉพาะเลน
  stopSpawnerForLane(laneIdx) {
    if (this.spawners && this.spawners[laneIdx]) {
      this.spawners[laneIdx].stop();
    }
    // หยุดเสียงรถที่ยังเล่นอยู่ในเลนนั้น
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

  // Returns currently playing cars in a lane
  getActiveCarsInLane(laneIdx) {
    return this.activeCarsPerLane[laneIdx] || [];
  }
}

class CarSpawner {
  constructor(direction, audioMixer, density, onCarSpawn) {
    this.direction = direction; // 'l' or 'r'
    this.audioMixer = audioMixer;
    this.density = density;
    this.onCarSpawn = onCarSpawn;
    this.running = false;
  }
  start() {
    this.running = true;
    // เพิ่ม delay เริ่มต้นแบบสุ่ม 0-4 วินาที เพื่อให้รถแต่ละเลนมาไม่พร้อมกัน
    // แต่เลนแรก (laneIdx === 0) ให้เริ่มทันที
    const initialDelay = (this.laneIdx === 0) ? 0 : Math.random() * 4000;
    setTimeout(() => {
      if (this.running) this.spawnLoop();
    }, initialDelay);
  }
  stop() { this.running = false; }

  async spawnLoop() {
    while (this.running) {
      // ถ้าเลนนี้ถูกข้ามไปแล้ว ไม่ต้อง spawn รถ
      if (window.game && window.game.currentLaneIdx > this.laneIdx) {
        await new Promise(res => setTimeout(res, 500));
        continue;
      }
      const carFiles = AUDIO.car[this.direction];
      const fileIdx = Math.floor(Math.random() * carFiles.length);
      const file = carFiles[fileIdx];
      // Fire car sound via mixer, pass carObj
      const carObj = await this.audioMixer.playCar(file, this.direction, this.laneIdx);
      if (!this.running) {
        if (carObj.source) { try { carObj.source.stop(); } catch (e) { } }
        if (carObj.audio) { try { carObj.audio.pause(); } catch (e) { } }
        break;
      }
      if (this.onCarSpawn) this.onCarSpawn(carObj);
      // รอให้รถคันก่อนเล่นไปแล้ว 90% + gap สุ่มเล็กน้อย
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
    // For each active car in the lane, determine if in danger zone (40-60% of audio)
    for (const car of cars) {
      let relTime = (audioCtxNow - car.startedAt);
      const duration = car.duration || 0;
      if (!duration || relTime < 0) continue;
      const dangerStart = 0.30 * duration;
      const dangerEnd = 0.60 * duration;
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
    this.carObjects = []; // global list of all playing cars (for mixing, not per-lane)
  }

  getVolumeForLane(laneIdx) {
    const laneDist = Math.abs(laneIdx - this.currentLane);
    if (laneDist === 0) {
      return 1; // เลนปัจจุบัน
    } else if (laneIdx < this.currentLane) {
      return 0.80; // เลนที่ข้ามไปแล้ว
    } else {
      // เลนที่ยังไปไม่ถึง
      if (laneDist === 1) {
        return 0.30;
      } else {
        return 0.05;
      }
    }
  }

  setCurrentLane(idx) {
    this.currentLane = idx;
    // Optionally, adjust volumes of cars playing in different lanes
    this.updateVolumes();
  }

  updateVolumes() {
    // ปรับความดังของรถในเลนอื่นให้เบากว่าเดิม
    this.carObjects.forEach(car => {
      const volume = this.getVolumeForLane(car.laneIdx);
      if (car.audio && car.audio.volume !== undefined) car.audio.volume = volume;
      if (car.gainNode) car.gainNode.gain.value = volume;
    });
  }

  async playCar(file, direction, laneIdx = 0) {
    // ปรับความดังของรถในเลนอื่นให้เบากว่าเดิม
    const volume = this.getVolumeForLane(laneIdx);
    const carObj = await this.audio.play(file, { volume });
    // Store extra info for collision
    carObj.direction = direction;
    carObj.laneIdx = laneIdx;
    this.carObjects.push(carObj);
    // Clean up car after done
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

  // Play global (not panned) event sound
  playEvent(file, { volume = 1 } = {}) {
    return this.audio.play(file, { volume });
  }
}

// ----------- Instructions Player (audio or text-to-speech) -----------
function playInstructions() {
  // เล่นไฟล์ instructions.mp3 เท่านั้น ไม่ใช้ TTS
  const audio = new Audio(AUDIO.instructions);
  audio.play();
}

function updateInstructionsRegion(msg) {
  document.getElementById('instructions-region').innerText = msg;
}

function updateSRFeedback(msg) {
  const region = document.getElementById('screenreader-feedback');
  region.innerText = msg;
}

// ----------------- UI Handler & Game Logic -----------------
class CrossroadGame {
  dangerInterval = null;

  startDangerCheck() {
    if (this.dangerInterval) return;
    this.dangerInterval = setInterval(() => {
      if (!this.gameActive) return;
      // ตรวจเฉพาะตอนที่ผู้เล่นอยู่กลางถนน (กดยังไม่ครบ 3 ครั้ง)
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
  }

  pauseGame() {
    if (!this.gameActive || this.isPaused) return;
    this.gameActive = false;
    this.isPaused = true;

    // รีเซ็ต clickCounter และหยุด dangerInterval ทันที
    this.clickCounter = 0;
    this.stopDangerCheck();

    // หยุดการปล่อยรถและเสียงรถทั้งหมด
    this.laneManager?.stopSpawners();
    this.audioMixer?.carObjects.forEach(car => {
      if (car.source) { try { car.source.stop(); } catch (e) { } }
      if (car.audio) { try { car.audio.pause(); } catch (e) { } }
    });
    if (this.audioMixer) this.audioMixer.carObjects = [];

    // ถอด role="application" และซ่อนปุ่มข้ามถนนชั่วคราว
    const crossBtn = document.getElementById('btn-cross');
    if (crossBtn) {
      crossBtn.removeAttribute('role');
      crossBtn.style.display = 'none';
    }

    // คืนค่าปุ่มสลับภาษา
    const langBtn = document.getElementById('btn-lang-toggle');
    if (langBtn) langBtn.style.display = '';

    // เปลี่ยนปุ่มเริ่มเกมเป็น "เล่นต่อ" และย้ายโฟกัสไปที่ปุ่มนั้นเพื่อออกจาก Focus Mode
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

    // ซ่อนปุ่มสลับภาษาและปุ่มเล่นต่อ
    const langBtn = document.getElementById('btn-lang-toggle');
    if (langBtn) langBtn.style.display = 'none';

    const startBtn = document.getElementById('btn-start-game');
    if (startBtn) startBtn.style.display = 'none';

    // คืนค่า role="application" และย้ายโฟกัสกลับไปที่ปุ่มข้ามถนน
    const crossBtn = document.getElementById('btn-cross');
    if (crossBtn) {
      crossBtn.style.display = '';
      crossBtn.setAttribute('role', 'application');
      crossBtn.focus();
    }

    // รีเซ็ต spawner ของด่านและเลนปัจจุบันเพื่อปล่อยรถใหม่ให้เวลาตรงกับ AudioContext
    this.laneManager?.stopSpawners();
    this.audioMixer = new AudioMixer(STAGES[this.stageIdx].lanes, this.currentLaneIdx);
    this.laneManager = new LaneManager(this.stageIdx);
    this.laneManager.startSpawners(this.audioMixer);
    this.audioMixer.setCurrentLane(this.currentLaneIdx);

    this.renderStatus();
  }

  async start() {
    // ซ่อนปุ่มสลับภาษาระหว่างเล่นเกม
    const langBtn = document.getElementById('btn-lang-toggle');
    if (langBtn) langBtn.style.display = 'none';

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

    // Initial instructions
    updateInstructionsRegion(t('instIntro'));
    playInstructions();
    this.stageIdx = 0;
    this.currentLaneIdx = 0;
    this.clickCounter = 0;
    this.gameActive = true;
    this.isPaused = false;
    this.lossState = false;
    this.winState = false;
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
      // ถ้าชน ให้รวมเลนที่ข้ามในด่านก่อนหน้า และเลนปัจจุบันที่ข้ามได้
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

    // แสดงผลบนหน้าจอด้วย
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
    this.showRestart();

    // แสดงปุ่มสลับภาษากลับคืนมาในหน้าจบเกม
    const langBtn = document.getElementById('btn-lang-toggle');
    if (langBtn) langBtn.style.display = '';

    // สรุปคะแนนรวมทุกด่าน
    this.updateEndResultUI(true);
  }

  showRestart() { document.getElementById('btn-restart').style.display = ''; }
  hideRestart() { document.getElementById('btn-restart').style.display = 'none'; }

  async handleCross() {
    if (!this.gameActive) return;

    this.clickCounter++;
    // Wait for required presses/taps
    if (this.clickCounter < CONTROLS_REQUIRED_CLICKS) {
      updateSRFeedback(t('pressMore')(CONTROLS_REQUIRED_CLICKS - this.clickCounter));
      // เริ่มตรวจสอบอันตรายเมื่อเริ่มข้าม (กดครั้งแรก)
      if (this.clickCounter === 1) {
        this.startDangerCheck();
      }
      return;
    }
    // หยุดตรวจสอบอันตรายเมื่อข้ามเลนสำเร็จหรือจบเกม
    this.stopDangerCheck();
    this.clickCounter = 0;

    // Get audio current time context
    const now = this.audioMixer.audio.ctx?.currentTime || (Date.now() / 1000);
    // Get cars playing in the current lane
    const carsInLane = this.laneManager.getActiveCarsInLane(this.currentLaneIdx);

    if (CollisionDetector.isCrossDanger(carsInLane, now)) {
      this.endGame(true);
      return;
    }
    // Safe crossing!
    this.audioMixer.playEvent(AUDIO.lanePass);
    // สรุปคะแนนรวมหลังข้ามแต่ละเลน
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
    // หยุดเสียงรถและการ spawn รถในเลนที่ข้ามไปแล้ว
    this.laneManager.stopSpawnerForLane(this.currentLaneIdx);
    this.currentLaneIdx++;
    if (this.currentLaneIdx >= STAGES[this.stageIdx].lanes) {
      // Stage finished
      this.stageIdx++;
      if (this.stageIdx >= STAGES.length) {
        // Win!
        this.endGame(false);
        return;
      }
      // Advance stage
      updateSRFeedback(t('stageReach')(this.stageIdx, STAGES[this.stageIdx].lanes));
      this.currentLaneIdx = 0;
      this.laneManager.stopSpawners();
      this.audioMixer = new AudioMixer(STAGES[this.stageIdx].lanes, this.currentLaneIdx);
      this.laneManager = new LaneManager(this.stageIdx);
      this.laneManager.startSpawners(this.audioMixer);
      // เริ่มตรวจสอบอันตรายเมื่อเริ่มเลนใหม่
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

  // Controls & event binding
  initControls() {
    // Keyboard: up arrow (3x) - ป้องกันการกดค้าง
    const pressedKeys = new Set();
    window.addEventListener('keydown', (e) => {
      // กด Escape เพื่อหยุดเกมชั่วคราวและออกจาก Focus Mode
      if (e.key === 'Escape' && this.gameActive) {
        e.preventDefault();
        this.pauseGame();
        return;
      }
      // กด R เพื่อ restart
      if ((e.key === 'r' || e.key === 'R') && !this.gameActive && (this.lossState || this.winState)) {
        updateSRFeedback(t('restarting'));
        this.start();
        return;
      }
      if (!this.gameActive) return;
      if (e.key === 'ArrowUp') {
        // ป้องกันการกดค้าง - ต้องปล่อยปุ่มก่อนจึงจะกดใหม่ได้
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

    // Mobile/desktop cross button
    const crossBtn = document.getElementById('btn-cross');
    if (crossBtn) {
      // ปุ่มจะจัดการ Enter/Space โดยอัตโนมัติให้กลายเป็น click event
      crossBtn.addEventListener('click', (e) => {
        if (this.gameActive) {
          this.handleCross();
        }
      });
    }

    // Restart
    document.getElementById('btn-restart').addEventListener('click', () => {
      updateSRFeedback(t('restarting'));
      this.start();
    });

    // Touch: also treat tap events
    document.getElementById('btn-cross').addEventListener('touchend', (e) => {
      e.preventDefault();
      this.handleCross();
    });

    // ARIA focus cues on controls
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
  // เรียก initControls หลัง DOM โหลดเสร็จ
  window.game.initControls();
  // ปุ่มเริ่มเล่นเกม / เล่นต่อ
  const startBtn = document.getElementById('btn-start-game');
  startBtn.addEventListener('click', () => {
    if (window.game.isPaused) {
      window.game.resumeGame();
    } else {
      window.game.start();
    }
  });
});
