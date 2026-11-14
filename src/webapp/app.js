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
    if (!age) { age = ageFromName(name.replace(/\.xlsx$/i, '')); if (age) ageSrc = 'ชื่อไฟล์'; }
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
    if (!files.includes(f)) call({ cmd: 'drop', id: f.id });
  } catch (e) {
    f.status = 'err';
    f.error = e;
  }
  render();
}

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
  go.disabled = busy || !ok.length || noAge.length > 0;
  $('goHint').textContent = busy ? 'กำลังอ่านไฟล์ กรุณารอสักครู่…'
    : !ok.length ? 'เลือกไฟล์อย่างน้อย 1 ไฟล์ และระบุกลุ่มอายุให้ครบทุกไฟล์'
    : noAge.length ? `ยังไม่ได้ระบุกลุ่มอายุ ${noAge.length} ไฟล์ กรุณาเลือกในช่องกลุ่มอายุ`
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
  if (o) $('lastInfo').textContent = `ปีงบประมาณ ${o.summary.fiscalYear} · ${fmt(o.summary.rowsTotal)} รายการ · ` +
    `${o.summary.units.length} หน่วยบริการ · สร้างเมื่อ ${o.summary.generatedAt}`;
}

function buildReport(summary) {
  const json = JSON.stringify(summary).replace(/</g, '\\u003c');
  let html = TEMPLATE.replace('/*__DATA__*/null', () => json);
  html = html.replace(/<head>/i, () => `<head>\n<meta http-equiv="Content-Security-Policy" content="${REPORT_CSP}">`);
  return html;
}

$('go').onclick = async () => {
  const go = $('go');
  go.disabled = true;
  go.textContent = 'กำลังคำนวณ…';
  showMsg('', '');
  try {
    const ok = files.filter(f => f.status === 'ok' && f.age);
    const res = await call({ cmd: 'summarise', files: ok.map(f => ({ id: f.id, age: f.age, name: f.name })),
                             hospitals: HOSP, version: APP.version, appName: APP.name });
    const s = res.summary;
    const sum = k => s.records.reduce((t, r) => t + r[k], 0);
    const ex = sum('examined'), other = sum('examinedOther');
    if (!ex) {
      showMsg('err', `ไม่พบผู้ที่ได้รับการตรวจฟันในปีงบประมาณ ${s.fiscalYear}`,
        [`มีผู้ได้รับการตรวจนอกปีงบประมาณนี้ ${fmt(other)} รายการ`, 'กรุณาตรวจสอบช่วงเวลาที่ส่งออกข้อมูลจาก HDC']);
      return;
    }
    const unknown = s.unknownCodes;
    delete s.unknownCodes;
    const info = `ปีงบประมาณ ${s.fiscalYear} · ${fmt(s.rowsTotal)} รายการ · ${s.units.length} หน่วยบริการ` +
      (unknown ? ` · ${unknown} รหัสไม่พบในตารางหน่วยบริการ (จัดเป็น "ไม่ทราบอำเภอ")` : '');
    const kept = saveLast(s, info);
    showReport(s, info + (kept ? '' : ' · (เบราว์เซอร์จำรายงานนี้ไว้ไม่ได้ กรุณากดบันทึกเป็นไฟล์)'));
  } catch (e) {
    showMsg('err', e.title || 'คำนวณไม่สำเร็จ', e.hints || [String(e)]);
  } finally {
    go.textContent = 'สร้าง Dashboard';
    render();
  }
};

function showReport(s, info) {
  lastSummary = s;
  lastHtml = buildReport(s);
  $('frame').srcdoc = lastHtml;
  $('vinfo').textContent = info;
  $('home').classList.add('hidden');
  $('viewer').classList.remove('hidden');
  document.title = `Dashboard สภาวะช่องปาก ปีงบประมาณ ${s.fiscalYear}`;
}
function openLast() {
  const o = loadLast();
  if (o) showReport(o.summary, `รายงานที่จำไว้ · สร้างเมื่อ ${o.summary.generatedAt} · ${o.info || ''}`);
}
$('openLast').onclick = openLast;
$('clearLast').onclick = () => { clearLast(); renderLastCard(); };

$('back').onclick = () => {
  renderLastCard();
  $('viewer').classList.add('hidden');
  $('home').classList.remove('hidden');
  $('frame').srcdoc = '';
  document.title = `Dashboard สภาวะช่องปาก (HDC Data Exchange) v${APP.version}`;
};

$('save').onclick = () => {
  if (!lastHtml) return;
  const d = new Date(), p = n => String(n).padStart(2, '0');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([lastHtml], { type: 'text/html;charset=utf-8' }));
  a.download = `Dashboardช่องปาก_ปีงบ${lastSummary.fiscalYear}_${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}.html`;
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
