#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
สร้างข้อมูลจำลองสำหรับทดสอบโปรแกรม (ไม่ใช่ข้อมูลจริง ไม่มีบุคคลจริง)

    python tools/make_sample_data.py [โฟลเดอร์ปลายทาง]

ค่าเริ่มต้นคือ sample-data/  โครงสร้างและชื่อคอลัมน์เหมือนไฟล์ที่ส่งออกจาก HDC
ใช้ในการทดสอบอัตโนมัติบน GitHub Actions และให้ผู้สนใจลองใช้โปรแกรมก่อนมีข้อมูลจริง
"""
import random
import sys
from datetime import date, timedelta
from pathlib import Path

import openpyxl

# หน้าจอของ Windows ใช้รหัสอักขระเดิม ทำให้พิมพ์ภาษาไทยแล้วโปรแกรมพัง จึงบังคับเป็น UTF-8
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

COLUMNS = ["hoscode", "hosname", "pid", "cid", "name", "lname", "sex", "birth", "addr",
           "check_vhid", "nation", "check_typearea", "discharge", "age_y", "seq",
           "date_serv", "denttype", "servplace", "pteeth", "pcaries", "pfilling",
           "pextract", "dteeth", "dcaries", "dfilling", "dextract", "need_fluoride",
           "need_scaling", "need_sealant", "need_pfilling", "need_dfilling",
           "need_pextract", "need_dextract", "nprosthesis", "permanent_permanent",
           "permanent_prosthesis", "prosthesis_prosthesis", "gum", "schooltype",
           "class", "provider", "providertype", "result"]

# รหัสหน่วยบริการสมมติ 6 แห่ง กระจายใน 2 อำเภอของจังหวัดตัวอย่าง
UNITS = [("03929", "รพ.สต.ตัวอย่างหนึ่ง"), ("03930", "รพ.สต.ตัวอย่างสอง"),
         ("03931", "รพ.สต.ตัวอย่างสาม"), ("04093", "รพ.สต.ตัวอย่างสี่"),
         ("10970", "โรงพยาบาลตัวอย่าง"), ("04007", "โรงพยาบาลตัวอย่างสอง")]

NA = "<NA>"
FY_START = date(2025, 10, 1)
FY_DAYS = 300
COUNTS = {3: 400, 6: 400, 12: 400, 60: 1200}


def make_row(rng, age, i):
    hoscode, hosname = rng.choice(UNITS)
    examined = rng.random() < 0.55                          # ~55% ได้รับการตรวจ
    row = {c: NA for c in COLUMNS}
    row.update({
        "hoscode": hoscode, "hosname": hosname,
        "pid": f"{i:06d}", "cid": "0" * 9 + "****",
        "name": "ตัวอย่าง", "lname": "ทดสอบ",
        "sex": rng.choice(["1", "2"]), "birth": "2000-01-01", "addr": "-",
        "check_vhid": "00000000", "nation": "099", "check_typearea": "1",
        "discharge": "9", "age_y": str(age), "result": "ตัวอย่าง",
    })
    if not examined:
        return [row[c] for c in COLUMNS]

    served = FY_START + timedelta(days=rng.randrange(FY_DAYS))
    row.update({
        "seq": str(100000 + i), "date_serv": served.isoformat(),
        "denttype": {3: "2", 6: "3", 12: "3", 60: "4"}[age],
        "servplace": rng.choice(["1", "2"]),
        "provider": "0001", "providertype": rng.choices(["06", "02", "05"], [80, 15, 5])[0],
        "need_fluoride": rng.choices(["1", "2"], [60, 40])[0],
        "need_scaling": rng.choices(["1", "2"], [25, 75])[0],
        "need_sealant": str(rng.choices([0, 1, 2], [80, 15, 5])[0]),
        "nprosthesis": "4", "permanent_permanent": "0",
        "permanent_prosthesis": "0", "prosthesis_prosthesis": "0",
        "pteeth": "0", "pcaries": "0", "pfilling": "0", "pextract": "0",
        "dteeth": "0", "dcaries": "0", "dfilling": "0", "dextract": "0",
        "need_pfilling": "0", "need_dfilling": "0",
        "need_pextract": "0", "need_dextract": "0",
        "gum": "000000",
    })

    if age == 3:
        dext = rng.choices([0, 1], [95, 5])[0]
        dcar = rng.choices([0, 1, 2, 4], [70, 12, 10, 8])[0]
        row.update({"dteeth": str(20 - dext), "dextract": str(dext), "dcaries": str(dcar),
                    "dfilling": str(rng.choices([0, 1], [92, 8])[0]),
                    "need_dfilling": str(dcar), "need_dextract": str(dext),
                    "class": "1", "gum": NA})
    elif age == 6:
        pteeth = rng.randrange(1, 13)
        pcar = rng.choices([0, 1], [95, 5])[0]
        row.update({"pteeth": str(pteeth), "pcaries": str(pcar),
                    "dteeth": str(rng.randrange(12, 21)),
                    "dcaries": str(rng.choices([0, 1, 3], [60, 20, 20])[0]),
                    "need_pfilling": str(pcar), "class": "2", "gum": NA})
    elif age == 12:
        pteeth = rng.randrange(20, 29)
        pcar = rng.choices([0, 1, 2], [75, 15, 10])[0]
        pfil = rng.choices([0, 1], [90, 10])[0]
        row.update({"pteeth": str(pteeth), "pcaries": str(pcar), "pfilling": str(pfil),
                    "pextract": "0", "need_pfilling": str(pcar), "class": "6",
                    "gum": rng.choices(["000000", "111111", "222222", "999999", "010000"],
                                       [70, 10, 12, 5, 3])[0]})
    else:
        pteeth = rng.randrange(0, 29)
        pext = 32 - pteeth
        occ = min(10, max(0, pteeth // 3))
        row.update({"pteeth": str(pteeth), "pextract": str(pext),
                    "pcaries": str(rng.choices([0, 1, 2], [70, 20, 10])[0]),
                    "pfilling": str(rng.choices([0, 1], [95, 5])[0]),
                    "permanent_permanent": str(occ),
                    "need_pextract": str(rng.choices([0, 1], [70, 30])[0]),
                    "gum": rng.choices(["000000", "111111", "222222", "333333",
                                        "999999", "444444"], [55, 10, 18, 8, 6, 3])[0]})
    return [row[c] for c in COLUMNS]


def main():
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "sample-data")
    rng = random.Random(20260811)                           # ผลลัพธ์เหมือนเดิมทุกครั้ง
    for age, count in COUNTS.items():
        folder = out / f"{age}ปี"
        folder.mkdir(parents=True, exist_ok=True)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(COLUMNS)
        for i in range(count):
            ws.append(make_row(rng, age, i))
        path = folder / f"sample_{age}y.xlsx"
        wb.save(path)
        print(f"สร้าง {path}  ({count:,} แถว)")


if __name__ == "__main__":
    main()
