#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
เติมชื่อ GitHub repository ลงในโค้ดก่อน build เพื่อให้โปรแกรมแจ้งเตือนเวอร์ชันใหม่ได้ถูกที่

    python packaging/set_repo.py                 หาจาก git remote ของเครื่องเอง
    python packaging/set_repo.py owner/repo      ระบุเอง

ถ้าหาไม่เจอจะไม่แก้อะไร โปรแกรมก็จะข้ามการตรวจสอบเวอร์ชันไปเฉย ๆ
"""
import re
import subprocess
import sys
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src" / "build_dashboard.py"
PATTERN = r'^GITHUB_REPO = ".*?"$'

# หน้าจอของ Windows ใช้รหัสอักขระเดิม (cp1252/cp874) ทำให้พิมพ์ภาษาไทยแล้วโปรแกรมพัง
# บังคับให้ output เป็น UTF-8 ก่อนเสมอ
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def from_git():
    try:
        url = subprocess.run(["git", "config", "--get", "remote.origin.url"],
                             capture_output=True, text=True, timeout=10).stdout.strip()
    except Exception:
        return None
    m = re.search(r"github\.com[:/]+([^/]+/[^/\s]+?)(?:\.git)?$", url)
    return m.group(1) if m else None


def main():
    repo = sys.argv[1] if len(sys.argv) > 1 else from_git()
    if not repo:
        print("ไม่พบชื่อ repository — ข้ามการตั้งค่าการแจ้งเตือนเวอร์ชันใหม่")
        return 0
    text = SRC.read_text(encoding="utf-8")
    if not re.search(PATTERN, text, flags=re.MULTILINE):
        print("ไม่พบบรรทัด GITHUB_REPO ในโค้ด")
        return 1
    new = re.sub(PATTERN, f'GITHUB_REPO = "{repo}"', text, count=1, flags=re.MULTILINE)
    if new == text:
        print(f"ตั้งค่าไว้เป็น {repo} อยู่แล้ว ไม่ต้องแก้")
        return 0
    SRC.write_text(new, encoding="utf-8")
    print(f"ตั้งค่าการแจ้งเตือนเวอร์ชันใหม่เป็น {repo}")
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
