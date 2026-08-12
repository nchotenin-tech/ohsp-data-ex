#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ตรวจว่าไฟล์ที่กำลังจะ commit ไม่มีข้อมูลผู้ป่วยติดไปด้วย
เรียกใช้โดย push_to_github.bat ก่อนส่งขึ้น GitHub

    python packaging/check_safe.py

คืนค่า 0 = ปลอดภัย, 1 = พบไฟล์ต้องสงสัย (สคริปต์เรียกใช้จะยกเลิกการส่ง)
"""
import subprocess
import sys

# รูปแบบชื่อไฟล์ที่ห้ามขึ้น GitHub เด็ดขาด
BLOCK_SUFFIX = (".xlsx", ".xls", ".xlsm")
BLOCK_NAMES = ("combined.csv", "summary.json", "meta.json", "log.txt")
BLOCK_DIRS = ("data/", "ผลลัพธ์/", "dist/", "build/")


def staged_files():
    out = subprocess.run(["git", "diff", "--cached", "--name-only", "-z"],
                         capture_output=True, timeout=60).stdout
    return [p.decode("utf-8", "replace") for p in out.split(b"\0") if p]


def main():
    files = staged_files()
    if not files:
        print("  ไม่มีไฟล์ที่เปลี่ยนแปลง")
        return 0

    bad = []
    for f in files:
        low = f.lower()
        if low.endswith(BLOCK_SUFFIX) or low.rsplit("/", 1)[-1] in BLOCK_NAMES \
                or any(f.startswith(d) or f"/{d}" in f for d in BLOCK_DIRS):
            bad.append(f)

    print(f"  ไฟล์ที่จะส่งขึ้น GitHub ทั้งหมด {len(files)} ไฟล์")
    for f in files:
        print(f"    {'!! ' if f in bad else '   '}{f}")

    if bad:
        print()
        print("  หยุด: พบไฟล์ที่อาจมีข้อมูลผู้ป่วย ยกเลิกการส่งเพื่อความปลอดภัย")
        print("  กรุณาตรวจสอบไฟล์ .gitignore แล้วลองใหม่")
        return 1
    print()
    print("  ผ่าน: ไม่พบไฟล์ข้อมูลผู้ป่วย")
    return 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    sys.exit(main())
