// เทียบผลคำนวณของเวอร์ชันเว็บ (src/webapp/core.js) กับเวอร์ชัน Python (summary.json)
//
//   node tools/parity_test.mjs <โฟลเดอร์ data> <summary.json จาก Python> <hospitals.json>
//
// โฟลเดอร์ data ต้องมีโครงสร้าง data/<กลุ่มอายุ>/*.xlsx แบบเดียวกับโปรแกรม .exe
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const OHSP = require('../src/webapp/core.js');
const [dataDir, refPath, hospPath] = process.argv.slice(2);

const files = [];
for (const folder of fs.readdirSync(dataDir).sort()) {
  const full = path.join(dataDir, folder);
  if (!fs.statSync(full).isDirectory()) continue;
  const age = OHSP.ageFromName(folder);
  if (!age) continue;
  for (const name of fs.readdirSync(full).sort()) {
    if (!name.toLowerCase().endsWith('.xlsx') || name.startsWith('~$')) continue;
    const buf = fs.readFileSync(path.join(full, name));
    const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const t0 = Date.now();
    const parsed = await OHSP.parseXlsx(ab);
    console.log(`  อ่าน ${folder}/${name}: ${parsed.rows.length} แถว (${Date.now() - t0} ms) เดาอายุจากข้อมูล = ${OHSP.ageFromData(parsed)}`);
    files.push({ age, name, parsed });
  }
}
files.sort((a, b) => OHSP.AGE_ORDER.indexOf(a.age) - OHSP.AGE_ORDER.indexOf(b.age) || (a.name < b.name ? -1 : 1));

const hosp = OHSP.expandHospitals(JSON.parse(fs.readFileSync(hospPath, 'utf8')));
const js = OHSP.summarise(files, hosp, 'test', 'test');
const py = JSON.parse(fs.readFileSync(refPath, 'utf8'));

let diffs = 0;
const report = (what, a, b) => { if (diffs++ < 30) console.log(`  ✗ ${what}: python=${JSON.stringify(a)} js=${JSON.stringify(b)}`); };
const day = s => String(s).slice(0, 10);
for (const k of ['rowsTotal', 'fiscalYear', 'fyStart', 'fyEnd']) if (py[k] !== js[k]) report(k, py[k], js[k]);
if (day(py.latestServe) !== day(js.latestServe)) report('latestServe', py.latestServe, js.latestServe);
if (JSON.stringify(py.units) !== JSON.stringify(js.units)) report('units', py.units.length, js.units.length);
if (py.records.length !== js.records.length) report('records.length', py.records.length, js.records.length);
let cells = 0;
py.records.forEach((r, i) => {
  const q = js.records[i] || {};
  for (const k in r) { cells++; if (r[k] !== q[k]) report(`records[${i}] ${r.age}/${r.hoscode}.${k}`, r[k], q[k]); }
});
console.log(diffs ? `\nไม่ตรงกัน ${diffs} จุด` : `\nตรงกันทุกค่า: ${py.records.length} ระเบียน ${cells} ค่า, ${py.units.length} หน่วยบริการ, ${py.rowsTotal} แถว`);
process.exit(diffs ? 1 : 0);
