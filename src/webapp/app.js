/* หน้าเลือกไฟล์และแสดง dashboard — ทำงานบนเธรดหลัก งานหนักส่งให้ Web Worker */
(function () {
'use strict';

const $ = id => document.getElementById(id);
const AGE_ORDER = ['3ปี', '6ปี', '12ปี', '60ปี'];
const AGE_TEXT = { '3ปี': '3 ปี', '6ปี': '6 ปี', '12ปี': '12 ปี', '60ปี': '60 ปีขึ้นไป' };
const AGE_COLOR = { '3ปี': 'var(--a3)', '6ปี': 'var(--a6)', '12ปี': 'var(--a12)', '60ปี': 'var(--a60)' };
const AGE_BY_NUMBER = { 3: '3ปี', 6: '6ปี', 12: '12ปี', 60: '60ปี' };
const fmt = n => Number(n).toLocaleString('th-TH');
const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

// รหัสรายงาน HDC ในชื่อไฟล์ที่ส่งออกมา
const REPORT_AGE = { '18.2': '3ปี', '18.3': '6ปี', '18.5': '12ปี', '18.11': '60ปี' };
function ageFromFileName(name) {
  const s = String(name || '');
  const code = s.match(/(?:^|[^\d.])(18\.(?:11|2|3|5))(?![\d.])/);
  if (code) return REPORT_AGE[code[1]];
  const m = s.match(/อายุ\s*(\d+)/);
  if (m && AGE_BY_NUMBER[+m[1]]) return AGE_BY_NUMBER[+m[1]];
  if (/ผู้สูงอายุ/.test(s)) return '60ปี';
  return ageFromName(s.replace(/\d+\.\d+/g, ' ').replace(/\(\d+\)/g, ' '));   // ตัดรหัสรายงาน (18.3) และเลขไฟล์ซ้ำ (2) ออกก่อน
}
function ageFromName(name) {
  const nums = String(name || '').match(/\d+/g) || [];
  for (const n of nums) if (AGE_BY_NUMBER[parseInt(n, 10)]) return AGE_BY_NUMBER[parseInt(n, 10)];
  return null;
}

/* ---------------- Worker (สำรองเป็นเธรดหลักถ้าเบราว์เซอร์ไม่อนุญาต) ---------------- */
const pending = new Map();
let worker = null, msgSeq = 0;
function startWorker() {
  const src = $('core-src').textContent;
  try {
    const url = URL.createObjectURL(new Blob([src], { type: 'text/javascript' }));
    worker = new Worker(url);
    worker.onmessage = ev => onWorkerMessage(ev.data);
    worker.onerror = ev => { ev.preventDefault(); fallbackToMainThread(); };
  } catch (e) {
    fallbackToMainThread();
  }
}
function fallbackToMainThread() {
  if (worker && worker.fake) return;
  if (worker && worker.terminate) try { worker.terminate(); } catch (e) { /* ignore */ }
  window.__OHSP_FAKE_WORKER__ = true;
  const s = document.createElement('script');
  s.textContent = $('core-src').textContent;
  document.head.appendChild(s);
  worker = {
    fake: true,
    postMessage: msg => setTimeout(() => window.__OHSP_HANDLE__(msg, m => onWorkerMessage(m)), 0),
  };
  // งานที่ค้างอยู่ใน worker เดิม ส่งใหม่
  for (const [id, p] of pending) if (p.msg) worker.postMessage(p.msg);
}
function call(msg, onProgress) {
  return new Promise((resolve, reject) => {
    const id = msg.id || ('m' + (++msgSeq));
    msg.id = id;
    pending.set(id, { resolve, reject, onProgress, msg });
    worker.postMessage(msg);
  });
}
function onWorkerMessage(m) {
  const p = pending.get(m.id);
  if (!p) return;
  if (m.type === 'progress') { if (p.onProgress) p.onProgress(m); return; }
  pending.delete(m.id);
  if (m.type === 'error') p.reject(m); else p.resolve(m);
}

/* ---------------- รายการไฟล์ ---------------- */
const files = [];          // {id, file, name, path, age, ageSrc, ageData, status, rows, error}
let fileSeq = 0, queue = Promise.resolve();
const notes = [];

function addFiles(list) {
  const skipped = [];
  for (const { file, path } of list) {
    const name = file.name;
    if (name.startsWith('~$')) { skipped.push(`${name} (ไฟล์ชั่วคราวของ Excel ที่เปิดค้างไว้)`); continue; }
    if (!/\.xlsx$/i.test(name)) {
      if (/\.(xls|csv|xlsm|ods)$/i.test(name)) skipped.push(`${name} (รองรับเฉพาะไฟล์ .xlsx ที่ส่งออกจาก HDC)`);
      continue;
    }
    if (files.some(f => f.name === name && f.file.size === file.size && f.file.lastModified === file.lastModified)) {
      skipped.push(`${name} (เลือกซ้ำ)`);
      continue;
    }
    // อายุจากชื่อโฟลเดอร์ที่อยู่ใกล้ไฟล์ที่สุด แล้วจึงดูชื่อไฟล์
    const dirs = (path || '').split('/').slice(0, -1).reverse();
    let age = null, ageSrc = null;
    for (const d of dirs) { age = ageFromName(d); if (age) { ageSrc = 'โฟลเดอร์ ' + d; break; } }
    if (!age) { age = ageFromFileName(name.replace(/\.xlsx$/i, '')); if (age) ageSrc = 'ชื่อไฟล์'; }
    const f = { id: 'f' + (++fileSeq), file, name, path: path || name, age, ageSrc, ageData: null,
                status: 'wait', rows: 0, error: null };
    files.push(f);
    queue = queue.then(() => parseFile(f));
  }
  if (skipped.length) notes.push({ kind: 'warn', title: 'ข้ามไฟล์ต่อไปนี้', items: skipped });
  render();
}

async function parseFile(f) {
  if (!files.includes(f)) return;
  f.status = 'reading';
  render();
  try {
    const res = await call({ cmd: 'parse', id: f.id, file: f.file }, m => { f.rows = m.rows; renderRow(f); });
    f.status = 'ok';
    f.rows = res.rows;
    f.ageData = res.ageData;
    if (!f.age && res.ageData) { f.age = res.ageData; f.ageSrc = 'คอลัมน์ age_y ในข้อมูล'; }
    // ชื่อไฟล์เชื่อถือได้น้อยกว่าข้อมูลจริง: ถ้าขัดกัน ใช้ค่าจากคอลัมน์ age_y (ชื่อโฟลเดอร์ผู้ใช้ตั้งเอง จึงคงไว้และแสดงป้ายเตือน)
    else if (f.ageSrc === 'ชื่อไฟล์' && res.ageData && res.ageData !== f.age) { f.age = res.ageData; f.ageSrc = 'คอลัมน์ age_y ในข้อมูล'; }
    if (!files.includes(f)) call({ cmd: 'drop', id: f.id });
  } catch (e) {
    f.status = 'err';
    f.error = e;
  }
  render();
}

/* ---------------- ช่วงเวลา ---------------- */
const TH_MON = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const thDate = s => { const [y, m, d] = s.split('-').map(Number); return `${d} ${TH_MON[m - 1]} ${y + 543}`; };
function getPeriod() {
  const custom = document.querySelector('input[name=per]:checked').value === 'custom';
  if (!custom) return { custom: false, ok: true };
  const start = $('pStart').value, end = $('pEnd').value;
  if (!start || !end) return { custom: true, ok: false, msg: 'กรุณาเลือกวันเริ่มต้นและวันสิ้นสุด' };
  if (start > end) return { custom: true, ok: false, msg: 'วันเริ่มต้นต้องไม่เกินวันสิ้นสุด' };
  return { custom: true, ok: true, start, end };
}
function renderPeriod() {
  const p = getPeriod(), h = $('pHint');
  h.classList.toggle('bad', !p.ok);
  h.textContent = !p.custom ? '' : !p.ok ? p.msg : `นับการตรวจตั้งแต่ ${thDate(p.start)} ถึง ${thDate(p.end)} (รวมวันสิ้นสุด)`;
}
document.querySelectorAll('input[name=per]').forEach(r => r.addEventListener('change', () => { renderPeriod(); render(); }));
['pStart', 'pEnd'].forEach(id => $(id).addEventListener('change', () => {
  document.querySelector('input[name=per][value=custom]').checked = true; renderPeriod(); render();
}));
const periodLabel = s => (s.customPeriod ? `ช่วง ${thDate(s.fyStart)} – ${thDate(s.fyEnd)}` : `ปีงบประมาณ ${s.fiscalYear}`);

function removeFile(id) {
  const i = files.findIndex(f => f.id === id);
  if (i < 0) return;
  const [f] = files.splice(i, 1);
  if (f.status === 'ok') call({ cmd: 'drop', id: f.id });
  render();
}

/* ---------------- วาดหน้าจอ ---------------- */
function statusHtml(f) {
  if (f.status === 'wait') return '<span class="st">รอคิว</span>';
  if (f.status === 'reading') return `<span class="st">กำลังอ่าน… ${f.rows ? fmt(f.rows) + ' แถว' : ''}<span class="bar"><i></i></span></span>`;
  if (f.status === 'ok') return `<span class="st ok">✓ ${fmt(f.rows)} แถว</span>`;
  const e = f.error || {};
  return `<span class="st err">✗ ${esc(e.title || 'อ่านไฟล์ไม่ได้')}</span>`;
}
function ageCell(f) {
  const opts = ['<option value="">— เลือก —</option>']
    .concat(AGE_ORDER.map(a => `<option value="${a}"${f.age === a ? ' selected' : ''}>${AGE_TEXT[a]}</option>`));
  let tag = '';
  if (f.status === 'ok' && f.ageData && f.age && f.ageData !== f.age)
    tag = `<div class="tag">ข้อมูลในไฟล์ดูเหมือน ${AGE_TEXT[f.ageData]}</div>`;
  else if (f.age && f.ageSrc && f.ageSrc.startsWith('คอลัมน์'))
    tag = `<div class="tag" style="background:#e0f2fe;color:#075985">เดาจากคอลัมน์ age_y</div>`;
  const need = f.status !== 'err' && !f.age ? ' need' : '';
  return `<select data-age="${f.id}" class="${need}"${f.status === 'err' ? ' disabled' : ''}>${opts.join('')}</select>${tag}`;
}
function rowHtml(f) {
  const sub = f.path !== f.name ? `<div class="path">${esc(f.path)}</div>` : '';
  const err = f.status === 'err' && f.error && f.error.hints && f.error.hints.length
    ? `<div class="path" style="color:var(--err)">${f.error.hints.map(esc).join('<br>')}</div>` : '';
  return `<td class="name">${esc(f.name)}${sub}${err}</td><td>${ageCell(f)}</td><td>${statusHtml(f)}</td>` +
         `<td><button class="x" data-rm="${f.id}" title="นำไฟล์นี้ออก">×</button></td>`;
}
function renderRow(f) {
  const tr = document.querySelector(`tr[data-id="${f.id}"]`);
  if (tr) tr.children[2].innerHTML = statusHtml(f);
}
function render() {
  const body = $('fileRows');
  body.innerHTML = files.map(f => `<tr data-id="${f.id}">${rowHtml(f)}</tr>`).join('');
  $('files').classList.toggle('hidden', !files.length);
  $('emptyList').classList.toggle('hidden', !!files.length);

  $('ages').innerHTML = AGE_ORDER.map(a => {
    const fs = files.filter(f => f.age === a && f.status === 'ok');
    const n = fs.reduce((s, f) => s + f.rows, 0);
    return `<div class="age${fs.length ? '' : ' empty'}"><div class="lab"><span class="dot" style="background:${AGE_COLOR[a]}"></span>อายุ ${AGE_TEXT[a]}</div>` +
           `<div class="n">${fs.length ? fmt(n) : '–'}</div><div class="f">${fs.length ? `แถว จาก ${fs.length} ไฟล์` : 'ยังไม่มีไฟล์'}</div></div>`;
  }).join('');

  $('notes').innerHTML = notes.map(n =>
    `<div class="msg ${n.kind}"><b>${esc(n.title)}</b><ul>${n.items.map(i => `<li>${esc(i)}</li>`).join('')}</ul></div>`).join('');

  const busy = files.some(f => f.status === 'wait' || f.status === 'reading');
  const ok = files.filter(f => f.status === 'ok');
  const noAge = ok.filter(f => !f.age);
  const go = $('go');
  const per = getPeriod();
  go.disabled = busy || !ok.length || noAge.length > 0 || !per.ok;
  $('goHint').textContent = busy ? 'กำลังอ่านไฟล์ กรุณารอสักครู่…'
    : !ok.length ? 'เลือกไฟล์อย่างน้อย 1 ไฟล์ และระบุกลุ่มอายุให้ครบทุกไฟล์'
    : noAge.length ? `ยังไม่ได้ระบุกลุ่มอายุ ${noAge.length} ไฟล์ กรุณาเลือกในช่องกลุ่มอายุ`
    : !per.ok ? `ช่วงเวลา: ${per.msg}`
    : `พร้อมแล้ว: ${ok.length} ไฟล์ รวม ${fmt(ok.reduce((s, f) => s + f.rows, 0))} แถว` +
      (files.some(f => f.status === 'err') ? ' (ไฟล์ที่อ่านไม่ได้จะไม่ถูกนำมาคำนวณ)' : '');
}

/* ---------------- การเลือกไฟล์ ---------------- */
$('pickFiles').onclick = () => $('inFiles').click();
$('pickFolder').onclick = () => $('inFolder').click();
$('inFiles').onchange = e => { addFiles([...e.target.files].map(file => ({ file, path: file.name }))); e.target.value = ''; };
$('inFolder').onchange = e => { addFiles([...e.target.files].map(file => ({ file, path: file.webkitRelativePath || file.name }))); e.target.value = ''; };

$('fileRows').addEventListener('click', e => { const id = e.target.dataset && e.target.dataset.rm; if (id) removeFile(id); });
$('fileRows').addEventListener('change', e => {
  const id = e.target.dataset && e.target.dataset.age;
  const f = files.find(x => x.id === id);
  if (f) { f.age = e.target.value || null; f.ageSrc = 'เลือกเอง'; render(); }
});

const drop = $('drop');
['dragenter', 'dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(t => drop.addEventListener(t, () => drop.classList.remove('over')));
// กันไม่ให้เบราว์เซอร์เปิดไฟล์แทนหน้านี้เมื่อวางพลาดนอกกรอบ
window.addEventListener('dragover', e => e.preventDefault());
window.addEventListener('drop', e => e.preventDefault());
drop.addEventListener('drop', async e => {
  e.preventDefault();
  const items = [...(e.dataTransfer.items || [])];
  const entries = items.map(i => i.webkitGetAsEntry && i.webkitGetAsEntry()).filter(Boolean);
  if (!entries.length) { addFiles([...e.dataTransfer.files].map(file => ({ file, path: file.name }))); return; }
  const out = [];
  const walk = async (entry, path) => {
    if (entry.isFile) {
      const file = await new Promise((res, rej) => entry.file(res, rej));
      out.push({ file, path: path + file.name });
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      for (;;) {
        const batch = await new Promise((res, rej) => reader.readEntries(res, rej));
        if (!batch.length) break;
        for (const child of batch) await walk(child, path + entry.name + '/');
      }
    }
  };
  for (const en of entries) await walk(en, '');
  addFiles(out);
});

/* ---------------- สร้าง dashboard ---------------- */
let lastHtml = null, lastSummary = null;
let ISSUES = null;                             // รายการไม่ผ่านเกณฑ์ (ไม่บันทึกลงเครื่องเด็ดขาด)
const HOSP = JSON.parse($('hosp').textContent);
const TEMPLATE = JSON.parse($('tpl').textContent);
const REPORT_CSP = "default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data: blob:; connect-src 'none'";

function showMsg(kind, title, items) {
  $('goMsg').innerHTML = title
    ? `<div class="msg ${kind}"><b>${esc(title)}</b>${items && items.length ? '<ul>' + items.map(i => `<li>${esc(i)}</li>`).join('') + '</ul>' : ''}</div>` : '';
}

/* ---------- จำรายงานล่าสุด (เฉพาะตัวเลขสรุประดับหน่วยบริการ) ---------- */
const LAST_KEY = 'ohsp-dashboard.lastReport.v1';
function loadLast() {
  try {
    const o = JSON.parse(localStorage.getItem(LAST_KEY) || 'null');
    return o && o.summary && Array.isArray(o.summary.records) ? o : null;
  } catch (e) { return null; }
}
function saveLast(summary, info) {
  try { localStorage.setItem(LAST_KEY, JSON.stringify({ summary, info })); return true; }
  catch (e) { return false; }
}
function clearLast() { try { localStorage.removeItem(LAST_KEY); } catch (e) { /* ignore */ } }
function renderLastCard() {
  const o = loadLast();
  $('lastCard').classList.toggle('hidden', !o);
  if (o) $('lastInfo').textContent = `${periodLabel(o.summary)} · ${fmt(o.summary.rowsTotal)} รายการ · ` +
    `${o.summary.units.length} หน่วยบริการ · สร้างเมื่อ ${o.summary.generatedAt}`;
}

function buildReport(summary) {
  const json = JSON.stringify(summary).replace(/</g, '\\u003c');
  let html = TEMPLATE.replace('/*__DATA__*/null', () => json);
  html = html.replace(/<head>/i, () => `<head>\n<meta http-equiv="Content-Security-Policy" content="${REPORT_CSP}">`);
  return html;
}

/** คำนวณจากไฟล์ที่อ่านไว้แล้ว คืน null ถ้าสำเร็จ หรือ {title, hints} ถ้าไม่สำเร็จ */
async function generate(period) {
  try {
    const ok = files.filter(f => f.status === 'ok' && f.age);
    if (!ok.length) return { title: 'ยังไม่มีไฟล์ข้อมูล', hints: ['กด “นำเข้าข้อมูลชุดใหม่” แล้วเลือกไฟล์ Excel'] };
    const res = await call({ cmd: 'summarise', files: ok.map(f => ({ id: f.id, age: f.age, name: f.name })),
                             hospitals: HOSP, version: APP.version, appName: APP.name, period });
    const s = res.summary;
    const sum = k => s.records.reduce((t, r) => t + r[k], 0);
    const ex = sum('examined'), other = sum('examinedOther');
    if (!ex) return { title: `ไม่พบผู้ที่ได้รับการตรวจฟันใน${periodLabel(s)}`,
                      hints: [`มีผู้ได้รับการตรวจนอกช่วงนี้ ${fmt(other)} รายการ`, 'กรุณาตรวจสอบช่วงเวลาที่เลือก หรือช่วงเวลาที่ส่งออกข้อมูลจาก HDC'] };
    ISSUES = res.issues || [];                 // มีข้อมูลระบุตัวบุคคล: เก็บในตัวแปรเท่านั้น
    const unknown = s.unknownCodes;
    delete s.unknownCodes;
    const info = `${periodLabel(s)} · ${fmt(s.rowsTotal)} รายการ · ${s.units.length} หน่วยบริการ` +
      (unknown ? ` · ${unknown} รหัสไม่พบในตารางหน่วยบริการ (จัดเป็น "ไม่ทราบอำเภอ")` : '');
    const kept = saveLast(s, info);
    showReport(s, info + (kept ? '' : ' · (เบราว์เซอร์จำรายงานนี้ไว้ไม่ได้ กรุณากดบันทึกเป็นไฟล์)'), true);
    return null;
  } catch (e) {
    return { title: e.title || 'คำนวณไม่สำเร็จ', hints: e.hints || [String(e)] };
  }
}
const periodArg = p => (p.custom && p.ok ? { start: p.start, end: p.end } : null);

$('go').onclick = async () => {
  const go = $('go');
  go.disabled = true;
  go.textContent = 'กำลังคำนวณ…';
  showMsg('', '');
  const err = await generate(periodArg(getPeriod()));
  if (err) showMsg('err', err.title, err.hints);
  go.textContent = 'สร้าง Dashboard';
  render();
};

/* ---------------- เปลี่ยนช่วงเวลาจากหน้ารายงาน ---------------- */
let CAN_RECALC = false, keepCov = null;
function perState() {
  const custom = document.querySelector('input[name=per2]:checked').value === 'custom';
  if (!custom) return { custom: false, ok: true };
  const start = $('p2Start').value, end = $('p2End').value;
  if (!start || !end) return { custom: true, ok: false, msg: 'กรุณาเลือกวันเริ่มต้นและวันสิ้นสุด' };
  if (start > end) return { custom: true, ok: false, msg: 'วันเริ่มต้นต้องไม่เกินวันสิ้นสุด' };
  return { custom: true, ok: true, start, end };
}
function renderPerPop() {
  const p = perState(), h = $('p2Hint');
  h.className = 'sub' + (p.ok ? '' : ' bad');
  h.textContent = !p.custom ? 'ใช้ปีงบประมาณล่าสุดที่พบในข้อมูล (1 ต.ค. – 30 ก.ย.)'
    : !p.ok ? p.msg : `นับการตรวจตั้งแต่ ${thDate(p.start)} ถึง ${thDate(p.end)} (รวมวันสิ้นสุด)`;
  $('p2Apply').disabled = !p.ok || !CAN_RECALC;
}
function openPerPop() {
  const s = lastSummary;
  document.querySelector(`input[name=per2][value=${s.customPeriod ? 'custom' : 'auto'}]`).checked = true;
  $('p2Start').value = s.fyStart.slice(0, 10); $('p2End').value = s.fyEnd.slice(0, 10);
  $('p2Msg').innerHTML = CAN_RECALC ? '' :
    '<div class="msg warn" style="margin:10px 0 0">รายงานนี้เป็นรายงานที่เบราว์เซอร์จำไว้ ไม่มีไฟล์ข้อมูลในหน่วยความจำ ' +
    'กด <b>นำเข้าข้อมูลชุดใหม่</b> เลือกไฟล์ แล้วกำหนดช่วงเวลาในขั้นที่ 2</div>';
  renderPerPop();
  $('perPop').classList.remove('hidden');
}
$('perBtn').onclick = () => ($('perPop').classList.contains('hidden') ? openPerPop() : $('perPop').classList.add('hidden'));
$('p2Cancel').onclick = () => $('perPop').classList.add('hidden');
document.querySelectorAll('input[name=per2]').forEach(r => r.addEventListener('change', renderPerPop));
['p2Start', 'p2End'].forEach(id => $(id).addEventListener('change', () => {
  document.querySelector('input[name=per2][value=custom]').checked = true; renderPerPop();
}));
$('p2Apply').onclick = async () => {
  const p = perState(); if (!p.ok) return;
  // ให้หน้าเลือกไฟล์ตรงกับที่เลือกที่นี่
  document.querySelector(`input[name=per][value=${p.custom ? 'custom' : 'auto'}]`).checked = true;
  if (p.custom) { $('pStart').value = p.start; $('pEnd').value = p.end; }
  renderPeriod();
  // คงเกณฑ์ร้อยละการตรวจที่ผู้ใช้เลือกไว้ในรายงานเดิม
  try { const d = $('frame').contentDocument; keepCov = { cov: d.getElementById('fCov').value, mode: (d.getElementById('fCovMode') || {}).value }; }
  catch (e) { keepCov = null; }
  const btn = $('p2Apply'); btn.disabled = true; btn.textContent = 'กำลังคำนวณ…';
  const err = await generate(periodArg(p));
  btn.textContent = 'คำนวณใหม่';
  if (err) { $('p2Msg').innerHTML = `<div class="msg err" style="margin:10px 0 0"><b>${esc(err.title)}</b>${(err.hints || []).map(esc).join('<br>')}</div>`; renderPerPop(); return; }
  $('perPop').classList.add('hidden');
};
$('frame').addEventListener('load', () => {
  if (keepCov == null) return;
  try {
    const d = $('frame').contentDocument;
    for (const [id, v] of [['fCovMode', keepCov.mode], ['fCov', keepCov.cov]]) {
      const sel = d.getElementById(id);
      if (sel && v != null && sel.value !== v) { sel.value = v; sel.dispatchEvent(new Event('change')); }
    }
  } catch (e) { /* ignore */ }
  keepCov = null;
});

function showReport(s, info, withIssues) {
  if (!withIssues) ISSUES = null;
  CAN_RECALC = !!withIssues;
  $('perPop').classList.add('hidden');
  $('perBtn').textContent = `ช่วงเวลา: ${periodLabel(s)} ▾`;
  const b = $('dqBtn');
  b.disabled = !ISSUES;
  b.textContent = ISSUES ? `ตรวจสอบคุณภาพข้อมูล (${fmt(ISSUES.length)})` : 'ตรวจสอบคุณภาพข้อมูล';
  b.title = ISSUES ? 'รายการที่ตรวจฟันในปีงบประมาณแต่ไม่ผ่านเงื่อนไขผู้ตรวจหรือจำนวนฟัน'
                   : 'รายงานที่จำไว้ไม่มีข้อมูลรายบุคคล กด “นำเข้าข้อมูลชุดใหม่” เพื่อดูรายการ';
  lastSummary = s;
  lastHtml = buildReport(s);
  $('frame').srcdoc = lastHtml;
  $('vinfo').textContent = info;
  $('home').classList.add('hidden');
  $('viewer').classList.remove('hidden');
  document.title = `Dashboard สภาวะช่องปาก ${periodLabel(s)}`;
}
function openLast() {
  const o = loadLast();
  if (o) showReport(o.summary, `รายงานที่จำไว้ · สร้างเมื่อ ${o.summary.generatedAt} · ${o.info || ''}`);
}
$('openLast').onclick = openLast;
$('clearLast').onclick = () => { clearLast(); renderLastCard(); };

$('back').onclick = () => {
  renderLastCard();
  $('perPop').classList.add('hidden');
  $('dqView').classList.add('hidden');
  $('viewer').classList.add('hidden');
  $('home').classList.remove('hidden');
  $('frame').srcdoc = '';
  document.title = `Dashboard สภาวะช่องปาก (HDC Data Exchange) v${APP.version}`;
};

/* ---------------- ตรวจสอบคุณภาพข้อมูล ---------------- */
const AGE_TXT2 = { '3ปี': '3 ปี', '6ปี': '6 ปี', '12ปี': '12 ปี', '60ปี': '60 ปีขึ้นไป' };
const MAX_ROWS = 500;
function dqFilter() {
  const age = $('dqAge').value, amp = $('dqAmp').value, hos = $('dqHos').value, why = $('dqWhy').value;
  const q = $('dqQ').value.trim().toLowerCase();
  return ISSUES.filter(r => (!age || r.age === age) && (!amp || r.ampname === amp) && (!hos || r.hoscode === hos)
    && (!why || (why === 'p' ? r.failProvider : r.failRange))
    && (!q || (r.name + ' ' + r.cid + ' ' + r.pid + ' ' + r.hosname).toLowerCase().includes(q)));
}
function fillSelect(el, items, allText, keep) {
  const cur = keep ? el.value : '';
  el.innerHTML = `<option value="">${allText}</option>` + items.map(([v, t]) => `<option value="${esc(v)}">${esc(t)}</option>`).join('');
  if (items.some(([v]) => v === cur)) el.value = cur;
}
function dqHosOptions() {
  const amp = $('dqAmp').value; const m = new Map();
  for (const r of ISSUES) if (!amp || r.ampname === amp) m.set(r.hoscode, `${r.hosname} (${r.hoscode})`);
  fillSelect($('dqHos'), [...m].sort((a, b) => (a[1] < b[1] ? -1 : 1)), 'ทุกหน่วยบริการ', true);
}
function dqRenderSummary() {
  const s = lastSummary, sum = (age, k) => s.records.filter(r => r.age === age).reduce((t, r) => t + r[k], 0);
  const rows = AGE_ORDER.map(a => {
    const ex = sum(a, 'examined'), fp = sum(a, 'failProvider'), fr = sum(a, 'failRange'), bad = fp + fr;
    return `<tr><td>อายุ ${AGE_TXT2[a]}</td><td>${fmt(ex)}</td><td class="${bad ? 'bad' : ''}">${fmt(bad)}</td>` +
      `<td>${ex ? (bad / ex * 100).toFixed(1) : '0.0'}%</td><td>${fmt(fp)}</td><td>${fmt(fr)}</td></tr>`;
  });
  $('dqSum').innerHTML = '<tr><th>กลุ่มอายุ</th><th>ตรวจในช่วงเวลา</th><th>ไม่ผ่านเกณฑ์</th><th>ร้อยละ</th>' +
    '<th>ผู้ตรวจไม่ใช่ทันตบุคลากร</th><th>จำนวนฟันไม่อยู่ในพิสัย</th></tr>' + rows.join('');
}
function dqRender() {
  const list = dqFilter();
  $('dqCount').textContent = list.length > MAX_ROWS
    ? `พบ ${fmt(list.length)} รายการ · แสดง ${fmt(MAX_ROWS)} รายการแรก (ไฟล์ CSV มีครบทุกรายการตามตัวกรอง)`
    : `พบ ${fmt(list.length)} รายการ`;
  $('dqCsv').disabled = !list.length;
  const head = '<tr><th>กลุ่มอายุ</th><th>อำเภอ</th><th>หน่วยบริการ</th><th>PID</th><th>SEQ</th><th>เลขบัตรประชาชน</th>' +
    '<th>ชื่อ-สกุล</th><th>วันที่ตรวจ</th><th>ผู้ตรวจ</th><th>providertype</th><th>ฟัน (ซี่/ผุ/อุด/ถอน)</th><th>สาเหตุที่ไม่ผ่าน</th></tr>';
  $('dqTable').innerHTML = head + list.slice(0, MAX_ROWS).map(r =>
    `<tr><td>${AGE_TXT2[r.age]}</td><td>${esc(r.ampname)}</td><td>${esc(r.hosname)}<div class="path">${esc(r.hoscode)}</div></td>` +
    `<td>${esc(r.pid)}</td><td>${esc(r.seq)}</td><td>${esc(r.cid)}</td><td>${esc(r.name)}</td><td>${esc(r.date)}</td>` +
    `<td>${esc(r.provider)}</td><td>${esc(r.providertype)}</td><td>${esc(r.teeth)}</td><td class="why">${esc(r.reason)}</td></tr>`).join('');
}
function openDq() {
  if (!ISSUES) return;
  fillSelect($('dqAge'), AGE_ORDER.filter(a => ISSUES.some(r => r.age === a)).map(a => [a, AGE_TXT2[a]]), 'ทุกกลุ่มอายุ');
  fillSelect($('dqAmp'), [...new Set(ISSUES.map(r => r.ampname))].sort().map(a => [a, a]), 'ทุกอำเภอ');
  dqHosOptions();
  $('dqWhy').value = ''; $('dqQ').value = '';
  $('dqInfo').textContent = `${periodLabel(lastSummary)} · ${thDate(lastSummary.fyStart)} ถึง ${thDate(lastSummary.fyEnd)}`;
  try {                                       // ใช้เงื่อนไขชุดเดียวกับในรายงาน (แหล่งเดียว)
    const src = $('frame').contentDocument.getElementById('rules1');
    $('dqRulesBody').innerHTML = src ? src.innerHTML : '';
    $('dqRules').classList.toggle('hidden', !src);
  } catch (e) { $('dqRules').classList.add('hidden'); }
  dqRenderSummary(); dqRender();
  $('dqView').classList.remove('hidden');
}
$('dqBtn').onclick = openDq;
$('dqBack').onclick = () => $('dqView').classList.add('hidden');
['dqAge', 'dqHos', 'dqWhy'].forEach(id => $(id).addEventListener('change', dqRender));
$('dqAmp').addEventListener('change', () => { dqHosOptions(); dqRender(); });
let dqT = 0;
$('dqQ').addEventListener('input', () => { clearTimeout(dqT); dqT = setTimeout(dqRender, 200); });
$('dqCsv').onclick = () => {
  const list = dqFilter();
  if (!list.length) return;
  const cols = [['age', 'กลุ่มอายุ', r => AGE_TXT2[r.age]], ['ampname', 'อำเภอ'], ['hoscode', 'รหัสหน่วยบริการ'], ['hosname', 'หน่วยบริการ'],
    ['pid', 'PID'], ['seq', 'SEQ'], ['cid', 'เลขบัตรประชาชน'], ['name', 'ชื่อ-สกุล'], ['date', 'วันที่ตรวจ'], ['provider', 'ผู้ตรวจ'],
    ['providertype', 'providertype'], ['teeth', 'ฟัน'], ['failProvider', 'ไม่ผ่าน_ผู้ตรวจ', r => (r.failProvider ? 1 : 0)],
    ['failRange', 'ไม่ผ่าน_จำนวนฟัน', r => (r.failRange ? 1 : 0)], ['reason', 'สาเหตุ']];
  // ใส่ ="..." ให้รหัสที่ขึ้นต้นด้วย 0 และเลข 13 หลัก ไม่ถูก Excel แปลงเป็นตัวเลข
  const keepText = new Set(['hoscode', 'pid', 'seq', 'cid', 'provider', 'providertype']);
  const cell = (k, v) => { v = String(v ?? ''); if (keepText.has(k) && /^\d+$/.test(v)) return `="${v}"`;
    return /[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v; };
  const lines = [cols.map(c => c[1]).join(',')].concat(list.map(r => cols.map(([k, , f]) => cell(k, f ? f(r) : r[k])).join(',')));
  const why = { p: 'ผู้ตรวจ', r: 'จำนวนฟัน' }[$('dqWhy').value];
  const scope = [$('dqAge').value, $('dqAmp').value, $('dqHos').value, why].filter(Boolean).join('_') || 'ทั้งหมด';
  const d = new Date(), p = n => String(n).padStart(2, '0');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }));
  a.download = `ตรวจสอบคุณภาพข้อมูล_${scope}_${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}.csv`;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
};

$('save').onclick = () => {
  if (!lastHtml) return;
  const d = new Date(), p = n => String(n).padStart(2, '0');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([lastHtml], { type: 'text/html;charset=utf-8' }));
  a.download = `Dashboardช่องปาก_${lastSummary.customPeriod ? lastSummary.fyStart + '_' + lastSummary.fyEnd : 'ปีงบ' + lastSummary.fiscalYear}_${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}.html`;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
};

function browserOk() {
  try { new DecompressionStream('deflate-raw'); new TextDecoderStream(); return true; } catch (e) { return false; }
}
if (!browserOk()) {
  $('drop').innerHTML = '<div class="msg err" style="text-align:left"><b>เบราว์เซอร์นี้เก่าเกินไป</b>' +
    'กรุณาเปิดไฟล์นี้ด้วย Google Chrome หรือ Microsoft Edge รุ่นปัจจุบัน (หรือ Firefox 113 ขึ้นไป)</div>';
} else {
  startWorker();
}
render();
renderLastCard();
if (loadLast()) openLast();                 // เปิดโปรแกรมแล้วเห็นรายงานล่าสุดทันที
})();
