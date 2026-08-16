#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ประกอบไฟล์ zip สำหรับแจกจ่าย
รันหลังจาก PyInstaller สร้าง dist/OralHealthDashboard.exe เรียบร้อยแล้ว

    python packaging/make_package.py [เวอร์ชัน]

ใช้ zipfile ของ Python เพื่อให้ชื่อไฟล์ภาษาไทยภายใน zip ถูกทำเครื่องหมาย UTF-8
(บางเครื่องมือบน Windows สร้าง zip ที่ทำให้ชื่อภาษาไทยเพี้ยนเมื่อแตกไฟล์)
"""
import re
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
# หน้าจอของ Windows ใช้รหัสอักขระเดิม ทำให้พิมพ์ภาษาไทยแล้วโปรแกรมพัง จึงบังคับเป็น UTF-8
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

EXE = ROOT / "dist" / "OralHealthDashboard.exe"

READ_ME_FIRST = """\
Dashboard วิเคราะห์สภาวะช่องปาก (HDC Data Exchange)
====================================================

วิธีใช้งาน

1. ดับเบิลคลิก OralHealthDashboard.exe หนึ่งครั้ง
   โปรแกรมจะสร้างโฟลเดอร์ data ให้อัตโนมัติ แล้วปิดตัวเอง

   หาก Windows แจ้งเตือน "Windows protected your PC"
   ให้กด More info แล้วกด Run anyway (ครั้งแรกครั้งเดียว)

2. นำไฟล์ Excel ที่ส่งออกจากระบบ HDC (Data Exchange การตรวจฟัน)
   มาวางในโฟลเดอร์ตามกลุ่มอายุ

       data\\3ปี\\      ข้อมูลกลุ่มอายุ 3 ปี
       data\\6ปี\\      ข้อมูลกลุ่มอายุ 6 ปี
       data\\12ปี\\     ข้อมูลกลุ่มอายุ 12 ปี
       data\\60ปี\\     ข้อมูลกลุ่มอายุ 60 ปีขึ้นไป

   วางได้มากกว่าหนึ่งไฟล์ต่อโฟลเดอร์ โปรแกรมจะนำมาต่อกันให้เอง
   กรุณาปิดไฟล์ใน Excel ให้เรียบร้อยก่อน

3. ดับเบิลคลิก OralHealthDashboard.exe อีกครั้ง
   รอประมาณ 1-3 นาที เสร็จแล้วเบราว์เซอร์จะเปิดรายงานขึ้นมาเอง

   ผลลัพธ์อยู่ที่  ผลลัพธ์\\dashboard.html


ข้อควรทราบเรื่องข้อมูลส่วนบุคคล

   ไฟล์ ผลลัพธ์\\dashboard.html  เป็นตัวเลขสรุประดับหน่วยบริการเท่านั้น
   ไม่มีข้อมูลรายบุคคล ส่งต่อให้ผู้อื่นได้อย่างปลอดภัย

   ไฟล์ในโฟลเดอร์ data มีข้อมูลรายบุคคลครบถ้วน
   กรุณาเก็บไว้ในเครื่องเท่านั้น


เมื่อเกิดปัญหา

   โปรแกรมจะแจ้งสาเหตุเป็นภาษาไทยบนหน้าจอ
   และบันทึกไว้ในไฟล์  ผลลัพธ์\\log.txt
   กรุณาส่งไฟล์นี้ให้ผู้ดูแลระบบเมื่อต้องการความช่วยเหลือ
"""


def main():
    version = sys.argv[1] if len(sys.argv) > 1 else read_version()
    if not EXE.exists():
        sys.exit(f"ไม่พบ {EXE} — กรุณารัน PyInstaller ก่อน")

    # วางไฟล์ไว้ที่ระดับบนสุดของ zip ไม่ต้องมีโฟลเดอร์ครอบอีกชั้น
    # เพราะ Windows สร้างโฟลเดอร์ตามชื่อไฟล์ zip ให้อยู่แล้วตอน Extract All
    # ผู้ใช้จึงเห็นโฟลเดอร์เดียวคือ OralHealthDashboard-vX.Y.Z แล้วเจอ .exe ทันที
    out = ROOT / f"OralHealthDashboard-{version}.zip"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(EXE, "OralHealthDashboard.exe")
        z.writestr("อ่านก่อนใช้งาน.txt", READ_ME_FIRST.encode("utf-8-sig"))
        for doc in sorted((ROOT / "docs").glob("*.docx")):
            z.write(doc, f"เอกสาร/{doc.name}")
        for doc in sorted((ROOT / "docs").glob("*.pdf")):
            z.write(doc, f"เอกสาร/{doc.name}")
    size = out.stat().st_size / 1024 / 1024
    print(f"สร้างแล้ว: {out.name}  ({size:.1f} MB)")
    with zipfile.ZipFile(out) as z:
        for n in z.namelist():
            print(f"   {n}")

    # แตกไฟล์เดียวกันไว้ในโฟลเดอร์ package/ ด้วย
    # GitHub Actions จะอัปโหลดโฟลเดอร์นี้เป็น Artifacts แล้วห่อ zip ให้เอง
    # ผู้ที่ดาวน์โหลดจาก Artifacts จึงเจอ .exe ทันทีเช่นกัน ไม่มีโฟลเดอร์ซ้อน
    stage = ROOT / "package"
    if stage.exists():
        shutil.rmtree(stage)
    with zipfile.ZipFile(out) as z:
        z.extractall(stage)
    print(f"เตรียมโฟลเดอร์ {stage.name}/ สำหรับอัปโหลดเป็น Artifacts แล้ว")


def read_version():
    src = (ROOT / "src" / "build_dashboard.py").read_text(encoding="utf-8")
    m = re.search(r'VERSION\s*=\s*"([^"]+)"', src)
    return "v" + (m.group(1) if m else "0.0.0")


if __name__ == "__main__":
    main()
