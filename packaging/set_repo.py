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
        return
    text = SRC.read_text(encoding="utf-8")
    new = re.sub(r'^GITHUB_REPO = ".*?"$', f'GITHUB_REPO = "{repo}"', text,
                 count=1, flags=re.MULTILINE)
    if new == text:
        print("ไม่พบบรรทัด GITHUB_REPO ในโค้ด")
        return
    SRC.write_text(new, encoding="utf-8")
    print(f"ตั้งค่าการแจ้งเตือนเวอร์ชันใหม่เป็น {repo}")


if __name__ == "__main__":
    main()
