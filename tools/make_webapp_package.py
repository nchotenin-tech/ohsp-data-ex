#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
สร้างไฟล์ zip พร้อมโครงสร้างโฟลเดอร์มาตรฐานสำหรับผู้ใช้

    python tools/make_webapp_package.py [โฟลเดอร์ปลายทาง]

ผลลัพธ์  Dashboard-HdcOralStatus-vX.Y.Z.zip

    Dashboard-HdcOralStatus\\
        OHSP-Dashboard-vX.Y.Z.html
        อ่านก่อนใช้งาน.txt
        data\\
            3ปี\\   6ปี\\   12ปี\\   60ปี\\      (มีไฟล์ .txt บอกว่าต้องวางรายงานไหน)
"""
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
import build_dashboard as bd  # noqa: E402

FOLDER = "Dashboard-HdcOralStatus"
AGES = [
    ("3ปี", "18.2 OHSP ร้อยละเด็กกลุ่มอายุ 3 ปีมีฟันผุในฟันน้ำนม"),
    ("6ปี", "18.3 OHSP ร้อยละเด็กกลุ่มอายุ 6 ปีมีฟันผุในฟันแท้"),
    ("12ปี", "18.5 OHSP ร้อยละเด็กกลุ่มอายุ 12 ปีมีฟันผุในฟันแท้"),
    ("60ปี", "18.11 OHSP ร้อยละของผู้สูงอายุ ที่มีฟันแท้ใช้งานได้ไม่น้อยกว่า 20 ซี่"),
]
CRLF = lambda s: s.replace("\r\n", "\n").replace("\n", "\r\n")


def readme(html_name):
    tree = "\n".join(f"        {a}\\        ← {r}" for a, r in AGES)
    return f"""Dashboard วิเคราะห์สภาวะช่องปาก (HDC Data Exchange) เวอร์ชัน {bd.VERSION}
==================================================================

โครงสร้างโฟลเดอร์มาตรฐาน
    {FOLDER}\\
        {html_name}    ← ดับเบิลคลิกเปิดด้วย Chrome หรือ Edge
        data\\
{tree}

วิธีใช้งาน
1. แตกไฟล์ zip นี้ไว้ที่ใดก็ได้ เช่น Documents (คลิกขวา > Extract All)
2. ส่งออกข้อมูลจาก HDC แบบ Data Exchange แล้ววางไฟล์ Excel ลงโฟลเดอร์ data ตามกลุ่มอายุ
   - ข้อมูลที่ HDC แบ่งเป็นหลายไฟล์ (ไฟล์ละ 100,000 รายการ) วางทุกไฟล์ในโฟลเดอร์เดียวกัน
   - ไม่ต้องเปลี่ยนชื่อไฟล์
   - ได้ข้อมูลชุดใหม่ ให้ลบหรือย้ายไฟล์ชุดเก่าออกจากโฟลเดอร์ก่อน (ไม่งั้นจะนับซ้ำ)
3. ดับเบิลคลิก {html_name}
4. กด "เลือกทั้งโฟลเดอร์" แล้วเลือกโฟลเดอร์ data
5. ตรวจช่องกลุ่มอายุ แล้วกด "สร้าง Dashboard"

ครั้งต่อไปที่เปิดไฟล์ HTML จะเห็นรายงานล่าสุดทันที
ต้องการส่งต่อรายงาน กด "บันทึกรายงานเป็นไฟล์" (มีแต่ตัวเลขสรุป ไม่มีข้อมูลรายบุคคล)

ข้อมูลส่วนบุคคล
- ไฟล์ Excel ในโฟลเดอร์ data มีเลขบัตรประชาชน ห้ามส่งต่อหรืออัปโหลดขึ้นที่สาธารณะ
- โปรแกรมอ่านและคำนวณในเครื่องนี้เท่านั้น หน้าเว็บถูกล็อกไม่ให้เชื่อมต่ออินเทอร์เน็ต

อัปเดตโปรแกรม: ดาวน์โหลดไฟล์ HTML เวอร์ชันใหม่จาก
    https://github.com/{bd.GITHUB_REPO}/releases
แล้ววางแทนไฟล์ HTML เดิมในโฟลเดอร์ {FOLDER} (โฟลเดอร์ data ใช้ต่อได้เลย)
"""


def main():
    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "dist"
    html = ROOT / "dist" / "OHSP-Dashboard.html"
    if not html.exists():
        raise SystemExit("ไม่พบ dist/OHSP-Dashboard.html — รัน tools/build_webapp.py ก่อน")
    html_name = f"OHSP-Dashboard-v{bd.VERSION}.html"
    target = out_dir / f"{FOLDER}-v{bd.VERSION}.zip"
    out_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"{FOLDER}/", "")
        z.write(html, f"{FOLDER}/{html_name}")
        z.writestr(f"{FOLDER}/อ่านก่อนใช้งาน.txt", "﻿" + CRLF(readme(html_name)))
        z.writestr(f"{FOLDER}/data/", "")
        for age, report in AGES:
            z.writestr(f"{FOLDER}/data/{age}/", "")
            z.writestr(f"{FOLDER}/data/{age}/วางไฟล์ Excel รายงาน {report.split()[0]} ที่นี่.txt",
                       "﻿" + CRLF(f"วางไฟล์ Excel ที่ส่งออกจาก HDC (Data Exchange) ของรายงาน\n{report}\n"
                                       f"ไว้ในโฟลเดอร์นี้ ถ้ามีหลายไฟล์ วางทุกไฟล์\n"))
    print(f"สร้าง {target} ({target.stat().st_size / 1024 / 1024:.2f} MB)")


if __name__ == "__main__":
    main()
