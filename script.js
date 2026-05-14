const currentTimeEl = document.getElementById('currentTime');
const alarmTimeInput = document.getElementById('alarmTime');
const setBtn = document.getElementById('setBtn');
const cancelBtn = document.getElementById('cancelBtn');
const statusEl = document.getElementById('status');
const alarmOverlay = document.getElementById('alarmOverlay');
const stopBtn = document.getElementById('stopBtn');

let alarmTarget = null;
let audioCtx = null;
let beepInterval = null;

function pad(n) {
  return String(n).padStart(2, '0');
}

function updateClock() {
  const now = new Date();
  currentTimeEl.textContent =
    `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;

  if (alarmTarget) {
    const h = now.getHours();
    const m = now.getMinutes();
    if (h === alarmTarget.h && m === alarmTarget.m && now.getSeconds() === 0) {
      triggerAlarm();
    }
  }
}

function beep() {
  if (!audioCtx) audioCtx = new (window.AudioContext || window.webkitAudioContext)();
  const osc = audioCtx.createOscillator();
  const gain = audioCtx.createGain();
  osc.connect(gain);
  gain.connect(audioCtx.destination);
  osc.frequency.value = 880;
  osc.type = 'sine';
  gain.gain.setValueAtTime(0.4, audioCtx.currentTime);
  gain.gain.exponentialRampToValueAtTime(0.001, audioCtx.currentTime + 0.5);
  osc.start();
  osc.stop(audioCtx.currentTime + 0.5);
}

function triggerAlarm() {
  alarmOverlay.classList.add('show');
  beepInterval = setInterval(beep, 800);
}

function stopAlarm() {
  alarmOverlay.classList.remove('show');
  clearInterval(beepInterval);
  beepInterval = null;
  alarmTarget = null;
  setBtn.disabled = false;
  cancelBtn.disabled = true;
  statusEl.textContent = 'アラームが設定されていません';
  statusEl.classList.remove('active');
}

setBtn.addEventListener('click', () => {
  const val = alarmTimeInput.value;
  if (!val) {
    statusEl.textContent = '時刻を選択してください';
    return;
  }
  const [h, m] = val.split(':').map(Number);
  alarmTarget = { h, m };
  setBtn.disabled = true;
  cancelBtn.disabled = false;
  statusEl.textContent = `${pad(h)}:${pad(m)} にアラームをセットしました`;
  statusEl.classList.add('active');
});

cancelBtn.addEventListener('click', () => {
  alarmTarget = null;
  setBtn.disabled = false;
  cancelBtn.disabled = true;
  statusEl.textContent = 'アラームをキャンセルしました';
  statusEl.classList.remove('active');
});

stopBtn.addEventListener('click', stopAlarm);

setInterval(updateClock, 1000);
updateClock();
