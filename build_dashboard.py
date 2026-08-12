#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ตัวเรียกใช้งานแบบสั้น — โค้ดจริงอยู่ที่ src/build_dashboard.py

    python build_dashboard.py [--force] [--full-csv] [--no-open]
"""
import runpy
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
runpy.run_path(str(Path(__file__).resolve().parent / "src" / "build_dashboard.py"),
               run_name="__main__")
