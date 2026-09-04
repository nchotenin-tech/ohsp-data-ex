#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
เลื่อนเลขเวอร์ชันย่อยขึ้นหนึ่งขั้นใน src/build_dashboard.py

    python packaging/bump_version.py          1.1.3 -> 1.1.4
    python packaging/bump_version.py minor    1.1.3 -> 1.2.0
    python packaging/bump_version.py major    1.1.3 -> 2.0.0
"""
import re
import sys
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src" / "build_dashboard.py"
PATTERN = r'^VERSION = "(\d+)\.(\d+)\.(\d+)"$'

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def main():
    part = (sys.argv[1] if len(sys.argv) > 1 else "patch").lower()
    text = SRC.read_text(encoding="utf-8")
    m = re.search(PATTERN, text, flags=re.MULTILINE)
    if not m:
        print("ไม่พบบรรทัด VERSION ในรูปแบบ x.y.z")
        return 1

    major, minor, patch = (int(x) for x in m.groups())
    if part == "major":
        major, minor, patch = major + 1, 0, 0
    elif part == "minor":
        minor, patch = minor + 1, 0
    else:
        patch += 1

    new_ver = f"{major}.{minor}.{patch}"
    SRC.write_text(re.sub(PATTERN, f'VERSION = "{new_ver}"', text, count=1,
                          flags=re.MULTILINE), encoding="utf-8")
    print(f"เลื่อนเวอร์ชันจาก {'.'.join(m.groups())} เป็น {new_ver}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
