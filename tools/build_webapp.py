#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
สร้างเว็บแอปไฟล์เดียว  dist/OHSP-Dashboard.html

    python tools/build_webapp.py

ประกอบจาก
    src/webapp/shell.html          หน้าเลือกไฟล์ (ฝั่งผู้ใช้)
    src/webapp/app.js              ตัวควบคุมหน้าจอ
    src/webapp/core.js             อ่าน Excel + คำนวณ (ทำงานใน Web Worker)
    src/dashboard_template.html    หน้า dashboard เดิม (ไม่แก้ไข)
    src/hospitals.csv              ตารางรหัสหน่วยบริการ → อำเภอ/จังหวัด

ผลลัพธ์ทำงานได้โดยไม่ต้องต่ออินเทอร์เน็ต และห้ามเชื่อมต่อเครือข่ายด้วย Content-Security-Policy
"""
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import build_dashboard as bd  # noqa: E402  ใช้ฟังก์ชันอ่าน hospitals.csv ตัวเดียวกับเวอร์ชัน .exe

VERSION = bd.VERSION          # เลขเวอร์ชันเดียวกับ src/build_dashboard.py (release.bat อ่านจากที่นั่น)
APP_NAME = bd.APP_NAME


def pack_hospitals():
    bd.say = lambda *a, **k: None
    mapping = bd.load_hospitals()
    pv, amp, lines = [], [], []
    pv_i, amp_i = {}, {}
    for code in sorted(mapping):
        h = mapping[code]
        for name, lst, idx in ((h["pvname"], pv, pv_i), (h["ampname"], amp, amp_i)):
            if name not in idx:
                idx[name] = len(lst)
                lst.append(name)
        clean = lambda s: s.replace("\t", " ").replace("\n", " ")
        lines.append(f"{clean(code)}\t{clean(h['hosname'])}\t{amp_i[h['ampname']]}\t{pv_i[h['pvname']]}")
    return {"pv": pv, "amp": amp, "h": "\n".join(lines)}, len(mapping)


def script_safe(text):
    """โค้ด JS ที่ฝังใน <script> ต้องไม่มีลำดับอักขระที่ทำให้ HTML ปิดแท็กก่อนเวลา"""
    for bad in ("</script", "<!--", "<script"):
        if bad in text.lower():
            raise SystemExit(f"โค้ดมีข้อความ {bad!r} ซึ่งฝังใน <script> ไม่ได้")
    return text


def json_safe(obj):
    """JSON ที่ฝังใน <script>: แปลง < เป็น \\u003c ทั้งหมด จึงปลอดภัยเสมอ"""
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")


def main():
    shell = (SRC / "webapp" / "shell.html").read_text(encoding="utf-8")
    app_js = (SRC / "webapp" / "app.js").read_text(encoding="utf-8")
    core_js = (SRC / "webapp" / "core.js").read_text(encoding="utf-8")
    template = (SRC / "dashboard_template.html").read_text(encoding="utf-8")
    if "/*__DATA__*/null" not in template:
        raise SystemExit("dashboard_template.html ไม่มีตำแหน่ง /*__DATA__*/null")
    hosp, n = pack_hospitals()

    out = shell
    for key, val in {
        "__VERSION__": VERSION,
        "__APP_NAME__": APP_NAME,
        "__REPO__": bd.GITHUB_REPO,
    }.items():
        out = out.replace(key, val)
    blocks = {
        "/*__CORE_JS__*/": script_safe(core_js),
        "/*__APP_JS__*/": script_safe(app_js),
        "/*__TEMPLATE_JSON__*/": json_safe(template),
        "/*__HOSPITALS_JSON__*/": json_safe(hosp),
    }
    for key, val in blocks.items():
        if key not in out:
            raise SystemExit(f"shell.html ไม่มีตำแหน่ง {key}")
        out = out.replace(key, val)

    dist = ROOT / "dist"
    dist.mkdir(exist_ok=True)
    target = dist / "OHSP-Dashboard.html"
    target.write_text(out, encoding="utf-8")
    (dist / "hospitals.packed.json").write_text(
        json.dumps(hosp, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"สร้าง {target}  ({target.stat().st_size / 1024 / 1024:.2f} MB, หน่วยบริการ {n:,} แห่ง, เวอร์ชัน {VERSION})")


if __name__ == "__main__":
    main()
