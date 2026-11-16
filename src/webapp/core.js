/*
 * OHSP core — อ่านไฟล์ Excel (.xlsx) และคำนวณสรุปตารางที่ 1–7 ภายในเบราว์เซอร์
 * ---------------------------------------------------------------------------
 * แปลงมาจาก src/build_dashboard.py (v1.2.0) ให้ได้ผลเหมือนกันทุกตัวเลข
 *  - ไม่ใช้ไลบรารีภายนอก แตกไฟล์ zip ด้วย DecompressionStream ที่มากับเบราว์เซอร์
 *  - เก็บเฉพาะคอลัมน์ที่ใช้คำนวณ ไม่เก็บชื่อ สกุล เลขบัตรประชาชน ที่อยู่ วันเกิด
 *  - ทำงานใน Web Worker ได้ (และทำงานบนเธรดหลักได้ถ้าเบราว์เซอร์ไม่อนุญาต Worker)
 */
(function (root) {
'use strict';

const AGE_ORDER = ['3ปี', '6ปี', '12ปี', '60ปี'];
const AGE_LABEL = { '3ปี': '3', '6ปี': '6', '12ปี': '12', '60ปี': '60+' };
const AGE_BY_NUMBER = { 3: '3ปี', 6: '6ปี', 12: '12ปี', 60: '60ปี' };
const NA_VALUES = new Set(['', '<na>', 'na', 'nan', 'none', 'null', '-']);

const REQUIRED_COLS = [
  'hoscode', 'date_serv', 'providertype',
  'pteeth', 'pcaries', 'pfilling', 'pextract',
  'dteeth', 'dcaries', 'dfilling', 'dextract',
  'need_fluoride', 'need_scaling', 'need_sealant',
  'need_pfilling', 'need_dfilling', 'need_pextract', 'need_dextract',
  'permanent_permanent', 'permanent_prosthesis', 'prosthesis_prosthesis', 'gum',
];
// คอลัมน์ที่เก็บเพิ่ม: hosname (ชื่อสำรองเมื่อรหัสไม่อยู่ใน hospitals) และ age_y (ใช้เดากลุ่มอายุ)
// ไม่มีข้อมูลระบุตัวบุคคล
const KEEP_COLS = REQUIRED_COLS.concat(['hosname', 'age_y']);
// ตัวระบุบุคคล: เก็บเฉพาะแถวที่มีวันที่ตรวจ ใช้แสดงในหน้าตรวจสอบคุณภาพข้อมูลเท่านั้น
// (อยู่ในหน่วยความจำชั่วคราว ไม่ถูกใส่ใน summary / รายงานที่บันทึก / localStorage)
const ID_COLS = ['pid', 'seq', 'cid', 'name', 'lname', 'provider'];

/* ======================================================================
 * 1) อ่านไฟล์ zip
 * ==================================================================== */
function u16(b, o) { return b[o] | (b[o + 1] << 8); }
function u32(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0; }

function readZipDirectory(buf) {
  const b = new Uint8Array(buf);
  let eocd = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) {
    if (b[i] === 0x50 && b[i + 1] === 0x4b && b[i + 2] === 0x05 && b[i + 3] === 0x06) { eocd = i; break; }
  }
  if (eocd < 0) throw new UserError('ไฟล์นี้ไม่ใช่ไฟล์ Excel (.xlsx) หรือไฟล์เสียหาย');
  let count = u16(b, eocd + 10);
  let off = u32(b, eocd + 16);
  // ZIP64 (ไฟล์ใหญ่มาก)
  if (off === 0xffffffff || count === 0xffff) {
    const loc = eocd - 20;
    if (loc >= 0 && u32(b, loc) === 0x07064b50) {
      const z = Number(new DataView(buf).getBigUint64(loc + 8, true));
      count = Number(new DataView(buf).getBigUint64(z + 32, true));
      off = Number(new DataView(buf).getBigUint64(z + 48, true));
    }
  }
  const dec = new TextDecoder('utf-8');
  const entries = {};
  for (let n = 0; n < count; n++) {
    if (u32(b, off) !== 0x02014b50) break;
    const method = u16(b, off + 10);
    let csize = u32(b, off + 20), usize = u32(b, off + 24);
    const nlen = u16(b, off + 28), xlen = u16(b, off + 30), clen = u16(b, off + 32);
    let lho = u32(b, off + 42);
    const name = dec.decode(b.subarray(off + 46, off + 46 + nlen));
    if (csize === 0xffffffff || usize === 0xffffffff || lho === 0xffffffff) {
      const dv = new DataView(buf);
      let p = off + 46 + nlen, end = p + xlen;
      while (p + 4 <= end) {
        const id = u16(b, p), sz = u16(b, p + 2);
        if (id === 1) {
          let q = p + 4;
          if (usize === 0xffffffff) { usize = Number(dv.getBigUint64(q, true)); q += 8; }
          if (csize === 0xffffffff) { csize = Number(dv.getBigUint64(q, true)); q += 8; }
          if (lho === 0xffffffff) { lho = Number(dv.getBigUint64(q, true)); q += 8; }
        }
        p += 4 + sz;
      }
    }
    entries[name.replace(/^\/+/, '')] = { method, csize, lho };
    off += 46 + nlen + xlen + clen;
  }
  return { buf, b, entries };
}

function findEntry(zip, path) {
  if (zip.entries[path]) return zip.entries[path];
  const low = path.toLowerCase();
  for (const k in zip.entries) if (k.toLowerCase() === low) return zip.entries[k];
  return null;
}

/** คืน ReadableStream ของข้อความ (UTF-8) จากไฟล์ในซิป */
function entryTextStream(zip, entry) {
  const b = zip.b, h = entry.lho;
  if (u32(b, h) !== 0x04034b50) throw new UserError('ไฟล์ Excel เสียหาย (โครงสร้าง zip ไม่ถูกต้อง)');
  const start = h + 30 + u16(b, h + 26) + u16(b, h + 28);
  const data = new Blob([b.subarray(start, start + entry.csize)]);
  let s = data.stream();
  if (entry.method === 8) s = s.pipeThrough(new DecompressionStream('deflate-raw'));
  else if (entry.method !== 0) throw new UserError('ไฟล์ Excel ใช้การบีบอัดแบบที่ไม่รองรับ');
  return s.pipeThrough(new TextDecoderStream('utf-8'));
}

async function entryText(zip, path) {
  const e = findEntry(zip, path);
  if (!e) return null;
  const reader = entryTextStream(zip, e).getReader();
  const parts = [];
  for (;;) { const { done, value } = await reader.read(); if (done) break; parts.push(value); }
  return parts.join('');
}

/* ======================================================================
 * 2) XML
 * ==================================================================== */
const ENT = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
function unescapeXml(s) {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (m, e) => {
    if (e[0] === '#') {
      const cp = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return String.fromCodePoint(cp);
    }
    return ENT[e] !== undefined ? ENT[e] : m;
  });
}
const ATTR_RE = {};
function attr(tag, name) {
  const re = ATTR_RE[name] || (ATTR_RE[name] = new RegExp('\\s' + name + '\\s*=\\s*(?:"([^"]*)"|\'([^\']*)\')'));
  const m = re.exec(tag);
  return m ? unescapeXml(m[1] !== undefined ? m[1] : m[2]) : null;
}
// ตัด namespace prefix ออก เช่น <x:row> -> <row> (บางโปรแกรมเขียนแบบนี้)
function stripNs(xml) { return /<[a-zA-Z0-9_]+:(row|c|v|is|t|si|r|rPh|sheetData)[\s>/]/.test(xml) ? xml.replace(/<(\/?)[a-zA-Z0-9_]+:/g, '<$1') : xml; }

/** ข้อความของ <si> หรือ <is>: รวม <t> ทั้งหมด ไม่รวมคำอ่าน <rPh> (เหมือน openpyxl) */
function richText(inner) {
  inner = inner.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
  let out = '', m;
  const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g;
  while ((m = re.exec(inner))) if (m[1] !== undefined) out += m[1];
  return unescapeXml(out.replace(/\r\n?/g, '\n'));
}

async function readSharedStrings(zip, path) {
  const e = path && findEntry(zip, path);
  if (!e) return [];
  const out = [];
  const reader = entryTextStream(zip, e).getReader();
  let buf = '';
  for (;;) {
    const { done, value } = await reader.read();
    if (value) buf += value;
    let last = 0, m;
    const re = /<(?:\w+:)?si\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?si>)/g;
    while ((m = re.exec(buf))) { out.push(m[1] === undefined ? '' : richText(stripNs(m[1]))); last = re.lastIndex; }
    buf = buf.slice(last);
    if (done) break;
  }
  return out;
}

/* ---------- รูปแบบตัวเลขที่เป็นวันที่ (เทียบ openpyxl.styles.numbers.is_date_format) ---------- */
const BUILTIN_DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);
function isDateFormat(code) {
  if (!code) return false;
  let s = code.split(';')[0];
  s = s.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, '');
  return /[dmyhs]/i.test(s);
}
function readDateStyles(xml) {
  const dates = new Set();
  if (!xml) return dates;
  xml = stripNs(xml);
  const custom = {};
  let m;
  const reF = /<numFmt\b[^>]*>/g;
  while ((m = reF.exec(xml))) custom[+attr(m[0], 'numFmtId')] = attr(m[0], 'formatCode');
  const cx = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
  if (!cx) return dates;
  const reX = /<xf\b[^>]*>/g;
  let i = 0;
  while ((m = reX.exec(cx[1]))) {
    const id = +(attr(m[0], 'numFmtId') || 0);
    if (BUILTIN_DATE_IDS.has(id) || (custom[id] !== undefined && isDateFormat(custom[id]))) dates.add(i);
    i++;
  }
  return dates;
}

/* ======================================================================
 * 3) แปลงค่าให้ได้ข้อความเหมือน str(value) ของ Python
 * ==================================================================== */
const pad = (n, w = 2) => String(n).padStart(w, '0');
/** เลขวันที่ของ Excel -> 'YYYY-MM-DD HH:MM:SS' แบบ str(datetime) ของ Python */
function excelDate(serial, date1904) {
  let days = Math.floor(serial);
  let secs = Math.round((serial - days) * 86400);
  if (secs >= 86400) { days += 1; secs -= 86400; }
  let ms;
  if (date1904) ms = Date.UTC(1904, 0, 1) + days * 864e5;
  else {
    // Excel นับ 29 ก.พ. 1900 ที่ไม่มีจริง (openpyxl ทำแบบเดียวกัน)
    if (days < 60) days += 1;
    ms = Date.UTC(1899, 11, 30) + days * 864e5;
  }
  const d = new Date(ms + secs * 1000);
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())} ` +
         `${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;
}
/** repr(float) ของ Python */
function pyFloatStr(x) {
  if (Number.isNaN(x)) return 'nan';
  if (!Number.isFinite(x)) return x > 0 ? 'inf' : '-inf';
  if (Object.is(x, -0)) return '-0.0';
  const a = Math.abs(x);
  if (a >= 1e16 || (a !== 0 && a < 1e-4)) {
    let s = x.toExponential();                              // 1e+21, 1.5e-7
    s = s.replace(/e([+-])(\d)$/, 'e$10$2');               // Python ใช้เลขชี้กำลังอย่างน้อย 2 หลัก
    return s;
  }
  const s = String(x);
  return /[.e]/.test(s) ? s : s + '.0';
}
/** ค่าในเซลล์ตัวเลข -> ข้อความแบบ openpyxl (_cast_number) + str() */
function numberCellStr(v) {
  if (/[.eE]/.test(v)) return pyFloatStr(parseFloat(v));
  const t = v.trim();
  if (!/^[+-]?\d+$/.test(t)) return pyFloatStr(parseFloat(v));
  return BigInt(t).toString();
}

/** ใช้ข้อความชุดเดียวกันซ้ำ (ค่าในคอลัมน์ที่เก็บมีไม่กี่แบบ) ลดหน่วยความจำลงมาก */
const INTERN = new Map();
function intern(s) {
  let v = INTERN.get(s);
  if (v === undefined) {
    v = s.length > 200 ? s : Array.from(s).join('');      // คัดลอกเพื่อไม่ให้ค้างอ้างอิงข้อความ XML ทั้งก้อน
    if (INTERN.size < 200000) INTERN.set(v, v);
  }
  return v;
}

/* ======================================================================
 * 4) อ่านแผ่นงาน
 * ==================================================================== */
function colIndex(ref) {
  let n = 0;
  for (let i = 0; i < ref.length; i++) {
    const c = ref.charCodeAt(i);
    if (c >= 65 && c <= 90) n = n * 26 + (c - 64);
    else if (c >= 97 && c <= 122) n = n * 26 + (c - 96);
    else break;
  }
  return n - 1;
}

/**
 * อ่านไฟล์ .xlsx หนึ่งไฟล์ คืนเฉพาะคอลัมน์ที่ต้องใช้
 * @returns {{cols: string[], rows: string[][], header: string[], sheet: string}}
 */
async function parseXlsx(buf, onProgress) {
  const zip = readZipDirectory(buf);
  const wbXml = await entryText(zip, 'xl/workbook.xml');
  if (!wbXml) throw new UserError('ไฟล์นี้ไม่ใช่ไฟล์ Excel (.xlsx) หรือไฟล์เสียหาย');
  const wb = stripNs(wbXml);
  const date1904 = /<workbookPr\b[^>]*\bdate1904\s*=\s*["'](1|true)["']/.test(wb);
  const relsXml = (await entryText(zip, 'xl/_rels/workbook.xml.rels')) || '';
  const rels = {};
  let m;
  const reR = /<(?:\w+:)?Relationship\b[^>]*>/g;
  while ((m = reR.exec(relsXml))) rels[attr(m[0], 'Id')] = { target: attr(m[0], 'Target'), type: attr(m[0], 'Type') || '' };
  const resolve = t => (t[0] === '/' ? t.slice(1) : 'xl/' + t).replace(/[^/]+\/\.\.\//g, '');

  const sheets = [];
  const reS = /<sheet\b[^>]*>/g;
  while ((m = reS.exec(wb))) {
    const rid = attr(m[0], 'r:id') || attr(m[0], 'id');
    if (rels[rid]) sheets.push({ name: attr(m[0], 'name'), path: resolve(rels[rid].target) });
  }
  if (!sheets.length) throw new UserError('ไม่พบแผ่นงานในไฟล์ Excel');
  const sheet = sheets.find(s => s.name === 'Data') || sheets[0];

  let sstPath = null, stylesPath = null;
  for (const id in rels) {
    if (/\/sharedStrings$/.test(rels[id].type)) sstPath = resolve(rels[id].target);
    if (/\/styles$/.test(rels[id].type)) stylesPath = resolve(rels[id].target);
  }
  const sst = await readSharedStrings(zip, sstPath || 'xl/sharedStrings.xml');
  const dateStyles = readDateStyles(await entryText(zip, stylesPath || 'xl/styles.xml'));

  const entry = findEntry(zip, sheet.path);
  if (!entry) throw new UserError('ไฟล์ Excel เสียหาย (ไม่พบแผ่นงาน ' + sheet.name + ')');

  const cellValue = (tag, inner) => {
    const t = attr(tag, 't') || 'n';
    if (t === 'inlineStr') {
      const is = /<is\b[^>]*>([\s\S]*?)<\/is>/.exec(inner);
      return is ? richText(is[1]) : null;
    }
    const vm = /<v\b[^>]*>([\s\S]*?)<\/v>|<v\b[^>]*\/>/.exec(inner);
    if (!vm) return null;
    const v = vm[1] === undefined ? '' : vm[1];
    if (t === 's') { const s = sst[parseInt(v, 10)]; return s === undefined ? null : s; }
    if (t === 'str' || t === 'e') return unescapeXml(v);
    if (t === 'b') return parseInt(v, 10) ? 'True' : 'False';
    if (t === 'd') return unescapeXml(v).replace('T', ' ');
    if (v === '') return null;
    const s = +(attr(tag, 's') || 0);
    if (dateStyles.has(s)) return excelDate(parseFloat(v), date1904);
    return numberCellStr(v);
  };

  let header = null, picks = null, idPicks = null, dsAt = -1, width = 0;
  const rows = [];
  const reCell = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  const handleRow = inner => {
    const vals = [];
    let next = 0, cm, any = false;
    reCell.lastIndex = 0;
    while ((cm = reCell.exec(inner))) {
      const tag = ' ' + cm[1];
      const r = attr(tag, 'r');
      const ci = r ? colIndex(r) : next;
      next = ci + 1;
      const v = cm[2] === undefined ? null : cellValue(tag, cm[2]);
      if (v !== null) { vals[ci] = v; any = true; }
    }
    if (header === null) {
      header = [];
      for (let i = 0; i < vals.length; i++) header.push(vals[i] == null ? '' : String(vals[i]).trim().toLowerCase());
      width = header.length;
      const missing = REQUIRED_COLS.filter(c => header.indexOf(c) < 0);
      if (missing.length) throw new UserError('ไฟล์ไม่มีคอลัมน์ที่จำเป็น', [
        'คอลัมน์ที่ขาด: ' + missing.join(', '),
        'กรุณาตรวจสอบว่าเป็นไฟล์ที่ส่งออกจากระบบ HDC (Data Exchange การตรวจฟัน) และไม่ได้ถูกแก้ไขหัวตาราง']);
      picks = KEEP_COLS.map(c => header.indexOf(c));
      idPicks = ID_COLS.map(c => header.indexOf(c));
      dsAt = header.indexOf('date_serv');
      return;
    }
    if (!any) return;                                     // แถวว่างทั้งแถว
    const row = picks.map(i => (i < 0 || vals[i] == null ? '' : intern(vals[i])));
    const ds = vals[dsAt] == null ? '' : String(vals[dsAt]).trim();
    if (ds && !NA_VALUES.has(ds.toLowerCase()))
      row.push(idPicks.map(i => (i < 0 || vals[i] == null ? '' : Array.from(String(vals[i])).join(''))));
    rows.push(row);
  };

  const reader = entryTextStream(zip, entry).getReader();
  let buf2 = '', seen = 0, started = false;
  const reRow = /<row\b[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g;
  for (;;) {
    const { done, value } = await reader.read();
    if (value) { buf2 += value; seen += value.length; }
    if (!started) {
      const i = buf2.search(/<(?:\w+:)?sheetData\b/);
      if (i < 0 && !done) continue;
      started = true;
      if (i > 0) buf2 = buf2.slice(i);
    }
    const chunk = /<\w+:row\b/.test(buf2.slice(0, 4000)) ? stripNs(buf2) : buf2;
    let last = 0, rm;
    reRow.lastIndex = 0;
    while ((rm = reRow.exec(chunk))) {
      if (rm[1] !== undefined) handleRow(rm[1]);
      else if (header === null) handleRow('');
      last = reRow.lastIndex;
    }
    buf2 = chunk.slice(last);
    if (onProgress) onProgress(rows.length, seen);
    if (done) break;
  }
  if (header === null) throw new UserError('ไฟล์ว่าง ไม่มีข้อมูล');
  return { cols: KEEP_COLS.slice(), rows, header, sheet: sheet.name, width };
}

/* ======================================================================
 * 5) กลุ่มอายุ
 * ==================================================================== */
function ageFromName(name) {
  const nums = String(name || '').match(/\d+/g) || [];
  for (const n of nums) if (AGE_BY_NUMBER[parseInt(n, 10)]) return AGE_BY_NUMBER[parseInt(n, 10)];
  return null;
}
/** เดากลุ่มอายุจากคอลัมน์ age_y (ค่าที่พบบ่อยที่สุด) */
function ageFromData(parsed) {
  const iy = parsed.cols.indexOf('age_y');
  const cnt = {};
  for (const r of parsed.rows) {
    const v = num(r[iy]);
    if (v === null) continue;
    const a = v >= 60 ? '60ปี' : AGE_BY_NUMBER[Math.floor(v)];
    if (a) cnt[a] = (cnt[a] || 0) + 1;
  }
  let best = null, n = 0;
  for (const a in cnt) if (cnt[a] > n) { best = a; n = cnt[a]; }
  const total = parsed.rows.length;
  return best && n >= total * 0.5 ? best : null;
}

/* ======================================================================
 * 6) ตัวช่วยแปลงค่า (เหมือน build_dashboard.py)
 * ==================================================================== */
const THAI0 = 0x0E50;
function toAsciiDigits(s) { return s.replace(/[\u0E50-\u0E59]/g, c => String(c.charCodeAt(0) - THAI0)); }
const FLOAT_RE = /^[+-]?(?:(?:\d(?:_?\d)*)(?:\.(?:\d(?:_?\d)*)?)?|\.\d(?:_?\d)*)(?:[eE][+-]?\d(?:_?\d)*)?$/;
function pyFloat(s) {
  s = toAsciiDigits(s);
  if (FLOAT_RE.test(s)) return parseFloat(s.replace(/_/g, ''));
  const l = s.toLowerCase().replace(/^\+/, '');
  if (l === 'inf' || l === 'infinity') return Infinity;
  if (l === '-inf' || l === '-infinity') return -Infinity;
  if (/^[+-]?nan$/.test(l)) return NaN;
  return null;
}
function num(value) {
  if (value === null || value === undefined) return null;
  const s = String(value).trim();
  if (NA_VALUES.has(s.toLowerCase())) return null;
  return pyFloat(s);
}
const n0 = v => { const x = num(v); return x === null ? 0 : x; };

function cleanGum(raw) {
  if (raw === null || raw === undefined) return null;
  let s = String(raw).trim();
  if (NA_VALUES.has(s.toLowerCase())) return null;
  s = toAsciiDigits(s.split('.')[0]);
  if (!/^\d+$/.test(s)) return null;
  if (s.length === 1) s = s.repeat(6);
  if (s.length !== 6) return null;
  const out = [];
  for (const ch of s) {
    let d = +ch;
    if (d === 4 || d === 5) d = 3;
    if (![0, 1, 2, 3, 9].includes(d)) return null;
    out.push(d);
  }
  return out;
}
function gumStatus(sx) {
  const valid = sx.filter(d => d <= 3);
  return valid.length ? Math.max(...valid) : 9;
}
function newBucket() {
  return {
    total: 0, examined: 0, examinedOther: 0, quality: 0,
    cariesFree: 0, cariesUntreated: 0, extracted: 0, filled: 0, dmftAny: 0,
    sumTeeth: 0, sumCaries: 0, sumExtract: 0, sumFilling: 0,
    gumN: 0, gum0: 0, gum1: 0, gum2: 0, gum3: 0, gum9: 0, gumBad: 0,
    sxt0: 0, sxt1: 0, sxt2: 0, sxt3: 0, sxt9: 0,
    func20: 0, occl4: 0, func20occl4: 0,
    needFluoride: 0, needSealant: 0, needFilling: 0, needExtract: 0, needScaling: 0,
    failProvider: 0, failRange: 0,
  };
}
const FLOAT_FIELDS = ['sumTeeth', 'sumCaries', 'sumExtract', 'sumFilling'];

/** วันที่ตรวจ: ถ้าเป็นเซลล์วันที่ของ Excel ใช้เฉพาะ YYYY-MM-DD */
function dateText(v) {
  const s = String(v || '').trim();
  return /^\d{4}-\d\d-\d\d 00:00:00$/.test(s) ? s.slice(0, 10) : s;
}

/* ======================================================================
 * 7) ปีงบประมาณและสรุปผล
 * ==================================================================== */
/**
 * @param files    [{age, name, parsed}] เรียงตามลำดับกลุ่มอายุแล้วชื่อไฟล์
 * @param hospitals {hoscode: {hosname, ampname, pvname}}
 */
/** เหตุผลที่จำนวนฟันไม่ผ่านเกณฑ์ (ใช้เฉพาะแถวที่ไม่ผ่าน จึงไม่กระทบผลคำนวณ) */
function rangeReasons(age, v) {
  const t = x => (x === null ? 'ว่าง' : String(x));
  const out = [];
  if (age === '3ปี') {
    const miss = ['dteeth', 'dcaries', 'dfilling', 'dextract'].filter(k => v[k] === null);
    if (miss.length) return ['ข้อมูลฟันน้ำนมว่าง: ' + miss.join(', ')];
    if (!(1 <= v.dteeth && v.dteeth <= 20)) out.push(`dteeth = ${t(v.dteeth)} (ต้องอยู่ระหว่าง 1–20)`);
    const s = v.dcaries + v.dfilling + v.dextract;
    if (!(s <= 20)) out.push(`dcaries+dfilling+dextract = ${s} (ต้องไม่เกิน 20)`);
    if (v.dteeth + v.dextract !== 20) out.push(`dteeth+dextract = ${v.dteeth + v.dextract} (ต้องเท่ากับ 20)`);
  } else if (age === '6ปี' || age === '12ปี') {
    const limit = age === '6ปี' ? 12 : 28;
    const miss = ['pteeth', 'pcaries', 'pfilling', 'pextract'].filter(k => v[k] === null);
    if (miss.length) return ['ข้อมูลฟันแท้ว่าง: ' + miss.join(', ')];
    if (!(1 <= v.pteeth && v.pteeth <= limit)) out.push(`pteeth = ${t(v.pteeth)} (ต้องอยู่ระหว่าง 1–${limit})`);
    const s = v.pfilling + v.pextract + v.pcaries;
    if (!(s <= limit)) out.push(`pcaries+pfilling+pextract = ${s} (ต้องไม่เกิน ${limit})`);
    if (!(v.pcaries + v.pfilling <= v.pteeth)) out.push(`pcaries+pfilling = ${v.pcaries + v.pfilling} มากกว่า pteeth = ${t(v.pteeth)}`);
  } else {
    const miss = ['pteeth', 'pextract', 'permanent_permanent', 'permanent_prosthesis', 'prosthesis_prosthesis'].filter(k => v[k] === null);
    if (miss.length) return ['ข้อมูลว่าง: ' + miss.join(', ')];
    if (!(0 <= v.pteeth && v.pteeth <= 32)) out.push(`pteeth = ${t(v.pteeth)} (ต้องอยู่ระหว่าง 0–32)`);
    const s = v.permanent_permanent + v.permanent_prosthesis + v.prosthesis_prosthesis;
    if (!(0 <= s && s <= 10)) out.push(`คู่สบฟันรวม = ${s} (ต้องอยู่ระหว่าง 0–10)`);
    if (!(v.pteeth + v.pextract > 0)) out.push('pteeth+pextract = 0 (ไม่มีข้อมูลฟัน)');
  }
  return out.length ? out : ['จำนวนฟันไม่ผ่านเงื่อนไข'];
}

/**
 * @param issues ถ้าส่ง array มา จะเก็บรายการที่ตรวจในปีงบแต่ไม่ผ่านเกณฑ์คุณภาพลงไป (มีข้อมูลระบุตัวบุคคล)
 */
function summarise(files, hospitals, version, appName, issues, period) {
  const C = {};
  KEEP_COLS.forEach((c, i) => (C[c] = i));

  let latest = null;
  for (const f of files) for (const r of f.parsed.rows) {
    const d = dateText(r[C.date_serv]);
    if (!NA_VALUES.has(d.toLowerCase()) && (latest === null || d > latest)) latest = d;
  }
  if (latest === null) throw new UserError('ไม่พบผู้ที่ได้รับการตรวจฟันเลยในข้อมูลชุดนี้', [
    'ทุกรายการมีคอลัมน์ date_serv ว่าง หมายความว่ายังไม่มีใครได้รับการตรวจ',
    'กรุณาตรวจสอบว่าส่งออกข้อมูลจาก HDC ถูกช่วงเวลาหรือไม่']);
  let fyBE, fyStart, fyEnd;
  const custom = !!(period && period.start && period.end);
  if (custom) {
    // ช่วงเวลาที่ผู้ใช้กำหนดเอง (YYYY-MM-DD ค.ศ.) ปีงบที่แสดงคิดจากวันสิ้นสุด
    fyStart = period.start; fyEnd = period.end;
    const ey = parseInt(fyEnd.slice(0, 4), 10), em = parseInt(fyEnd.slice(5, 7), 10);
    fyBE = (em >= 10 ? ey + 1 : ey) + 543;
  } else {
    const year = parseInt(latest.slice(0, 4), 10), month = parseInt(latest.slice(5, 7), 10);
    const endYear = month >= 10 ? year + 1 : year;
    fyBE = endYear + 543; fyStart = `${endYear - 1}-10-01`; fyEnd = `${endYear}-09-30`;
  }

  const buckets = new Map();
  const units = {};
  const unknown = new Set();
  let rowsTotal = 0;

  for (const f of files) {
    const age = f.age;
    for (const row of f.parsed.rows) {
      rowsTotal++;
      if (!AGE_LABEL[age]) continue;
      const g = c => row[C[c]];
      const hoscode = String(g('hoscode')).trim();
      const key = age + '\u0000' + hoscode;
      let b = buckets.get(key);
      if (!b) { b = newBucket(); buckets.set(key, b); }
      b.total++;

      if (!(hoscode in units)) {
        let info = hospitals[hoscode];
        if (!info) { unknown.add(hoscode); info = {}; }
        units[hoscode] = {
          hoscode,
          hosname: info.hosname || String(g('hosname')).trim() || hoscode,
          ampname: info.ampname || 'ไม่ทราบอำเภอ',
          pvname: info.pvname || 'ไม่ทราบจังหวัด',
        };
      }

      const ds = dateText(g('date_serv'));
      if (NA_VALUES.has(ds.toLowerCase())) continue;
      if (!(fyStart <= ds && ds <= fyEnd)) { b.examinedOther++; continue; }
      b.examined++;

      // ---------- เกณฑ์คุณภาพ ----------
      const providerOk = ['02', '06'].includes(String(g('providertype')).trim());
      const pteeth = num(g('pteeth')), pcaries = num(g('pcaries'));
      const pfilling = num(g('pfilling')), pextract = num(g('pextract'));
      const dteeth = num(g('dteeth')), dcaries = num(g('dcaries'));
      const dfilling = num(g('dfilling')), dextract = num(g('dextract'));
      const pp = num(g('permanent_permanent')), ppr = num(g('permanent_prosthesis'));
      const prpr = num(g('prosthesis_prosthesis'));
      const has = (...a) => a.every(x => x !== null);

      let rangeOk;
      if (age === '3ปี') {
        rangeOk = has(dteeth, dcaries, dfilling, dextract)
          && 1 <= dteeth && dteeth <= 20
          && dcaries + dfilling + dextract <= 20
          && dteeth + dextract === 20;
      } else if (age === '6ปี' || age === '12ปี') {
        const limit = age === '6ปี' ? 12 : 28;
        rangeOk = has(pteeth, pcaries, pfilling, pextract)
          && 1 <= pteeth && pteeth <= limit
          && pfilling + pextract + pcaries <= limit
          && pcaries + pfilling <= pteeth;
      } else {
        const s = pp + ppr + prpr;
        rangeOk = has(pteeth, pextract, pp, ppr, prpr)
          && 0 <= pteeth && pteeth <= 32
          && 0 <= s && s <= 10
          && pteeth + pextract > 0;
      }
      if (!providerOk) b.failProvider++;
      else if (!rangeOk) b.failRange++;
      if (!(providerOk && rangeOk)) {
        if (issues) {
          const id = Array.isArray(row[row.length - 1]) ? row[row.length - 1] : [];
          const reasons = [];
          if (!providerOk) reasons.push(`providertype = ${String(g('providertype')).trim() || 'ว่าง'} (ต้องเป็น 02 หรือ 06)`);
          if (!rangeOk) reasons.push(...rangeReasons(age, { dteeth, dcaries, dfilling, dextract, pteeth, pcaries, pfilling, pextract,
            permanent_permanent: pp, permanent_prosthesis: ppr, prosthesis_prosthesis: prpr }));
          const u = units[hoscode];
          const teeth = age === '3ปี'
            ? `d ${String(g('dteeth'))}/${String(g('dcaries'))}/${String(g('dfilling'))}/${String(g('dextract'))}`
            : `p ${String(g('pteeth'))}/${String(g('pcaries'))}/${String(g('pfilling'))}/${String(g('pextract'))}`
              + (age === '60ปี' ? ` · คู่สบ ${String(g('permanent_permanent'))}/${String(g('permanent_prosthesis'))}/${String(g('prosthesis_prosthesis'))}` : '');
          issues.push({ age, ampname: u.ampname, hoscode, hosname: u.hosname,
            pid: id[0] || '', seq: id[1] || '', cid: id[2] || '', name: [id[3], id[4]].filter(Boolean).join(' '),
            provider: id[5] || '', providertype: String(g('providertype')).trim(), date: ds, teeth,
            failProvider: !providerOk, failRange: !rangeOk, reason: reasons.join(' ; ') });
        }
        continue;
      }
      b.quality++;

      // ---------- ตารางที่ 2 และ 3 ----------
      let teeth, caries, filling, extract;
      if (age === '3ปี') { teeth = n0(dteeth); caries = n0(dcaries); filling = n0(dfilling); extract = n0(dextract); }
      else { teeth = n0(pteeth); caries = n0(pcaries); filling = n0(pfilling); extract = n0(pextract); }
      if (caries + filling + extract === 0) b.cariesFree++; else b.dmftAny++;
      b.cariesUntreated += caries > 0 ? 1 : 0;
      b.extracted += extract > 0 ? 1 : 0;
      b.filled += filling > 0 ? 1 : 0;
      b.sumTeeth += teeth; b.sumCaries += caries; b.sumExtract += extract; b.sumFilling += filling;

      // ---------- ตารางที่ 4 และ 5 ----------
      if (age === '12ปี' || age === '60ปี') {
        const sx = cleanGum(g('gum'));
        if (sx === null) b.gumBad++;
        else {
          b.gumN++;
          b['gum' + gumStatus(sx)]++;
          for (const d of sx) b['sxt' + d]++;
        }
      }

      // ---------- ตารางที่ 6 ----------
      if (age === '60ปี') {
        const f20 = n0(pteeth) >= 20;
        const o4 = n0(pp) + n0(ppr) + n0(prpr) >= 4;
        b.func20 += f20 ? 1 : 0;
        b.occl4 += o4 ? 1 : 0;
        b.func20occl4 += f20 && o4 ? 1 : 0;
      }

      // ---------- ตารางที่ 7 ----------
      b.needFluoride += String(g('need_fluoride')).trim() === '1' ? 1 : 0;
      b.needScaling += String(g('need_scaling')).trim() === '1' ? 1 : 0;
      b.needSealant += n0(g('need_sealant')) > 0 ? 1 : 0;
      if (age === '3ปี') {
        b.needFilling += n0(g('need_dfilling')) > 0 ? 1 : 0;
        b.needExtract += n0(g('need_dextract')) > 0 ? 1 : 0;
      } else {
        b.needFilling += n0(g('need_pfilling')) > 0 ? 1 : 0;
        b.needExtract += n0(g('need_pextract')) > 0 ? 1 : 0;
      }
    }
  }

  const cmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
  const keys = [...buckets.keys()].sort((x, y) => {
    const [a1, h1] = x.split('\u0000'), [a2, h2] = y.split('\u0000');
    return cmp(a1, a2) || cmp(h1, h2);
  });
  const records = keys.map(k => {
    const [age, hoscode] = k.split('\u0000');
    const b = buckets.get(k);
    const rec = { age, hoscode };
    for (const f in b) rec[f] = FLOAT_FIELDS.includes(f) ? Math.round(b[f] * 100) / 100 : b[f];
    return rec;
  });

  const now = new Date();
  return {
    app: appName,
    version,
    generatedAt: `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}`,
    rowsTotal,
    fiscalYear: fyBE,
    ...(custom ? { customPeriod: true } : {}),
    fyStart,
    fyEnd,
    latestServe: latest,
    ageOrder: AGE_ORDER,
    ageLabel: AGE_LABEL,
    units: Object.keys(units).sort(cmp).map(k => units[k]),
    records,
    unknownCodes: unknown.size,
  };
}

/** แปลงตาราง hospitals ที่ฝังมาแบบย่อให้เป็น {hoscode: {...}} */
function expandHospitals(packed) {
  const out = {};
  for (const line of packed.h.split('\n')) {
    if (!line) continue;
    const [code, name, ai, pi] = line.split('\t');
    out[code] = { hosname: name, ampname: packed.amp[+ai] || '', pvname: packed.pv[+pi] || '' };
  }
  return out;
}

class UserError extends Error {
  constructor(title, hints) { super(title); this.title = title; this.hints = hints || []; this.isUserError = true; }
}

const OHSP = {
  AGE_ORDER, AGE_LABEL, REQUIRED_COLS, KEEP_COLS, ID_COLS,
  parseXlsx, summarise, ageFromName, ageFromData, expandHospitals, UserError,
  _t: { num, cleanGum, pyFloatStr, numberCellStr, excelDate, isDateFormat, dateText },
};
root.OHSP = OHSP;
if (typeof module !== 'undefined' && module.exports) module.exports = OHSP;

/* ======================================================================
 * 8) ตัวรับคำสั่งเมื่อทำงานใน Web Worker
 * ==================================================================== */
const inWorker = typeof WorkerGlobalScope !== 'undefined' && root instanceof WorkerGlobalScope;
if (inWorker || root.__OHSP_FAKE_WORKER__) {
  const store = new Map();                                 // id -> parsed (อยู่ในหน่วยความจำเท่านั้น)
  let HOSP = null;
  const handle = async (msg, post) => {
    const { cmd, id } = msg;
    try {
      if (cmd === 'parse') {
        let lastTick = 0;
        const buf = msg.buffer || await msg.file.arrayBuffer();
        const parsed = await parseXlsx(buf, (rows) => {
          const t = Date.now();
          if (t - lastTick > 250) { lastTick = t; post({ id, type: 'progress', rows }); }
        });
        store.set(id, parsed);
        post({ id, type: 'parsed', rows: parsed.rows.length, sheet: parsed.sheet,
               ageData: ageFromData(parsed), hasAgeY: parsed.header.includes('age_y') });
      } else if (cmd === 'drop') {
        store.delete(id);
      } else if (cmd === 'clear') {
        store.clear();
      } else if (cmd === 'summarise') {
        if (!HOSP) HOSP = expandHospitals(msg.hospitals);
        const order = (a, b) => AGE_ORDER.indexOf(a.age) - AGE_ORDER.indexOf(b.age) || (a.name < b.name ? -1 : a.name > b.name ? 1 : 0);
        const files = msg.files.filter(f => store.has(f.id)).map(f => ({ ...f, parsed: store.get(f.id) })).sort(order);
        const issues = [];
        const summary = summarise(files, HOSP, msg.version, msg.appName, issues, msg.period);
        post({ id, type: 'summary', summary, issues });
      }
    } catch (e) {
      post({ id, type: 'error', title: e.isUserError ? e.title : 'เกิดข้อผิดพลาดที่ไม่คาดคิด',
             hints: e.isUserError ? e.hints : [String(e && e.message || e)], user: !!e.isUserError });
    }
  };
  if (inWorker) root.onmessage = ev => handle(ev.data, m => root.postMessage(m));
  else root.__OHSP_HANDLE__ = handle;
}
})(typeof self !== 'undefined' ? self : globalThis);
