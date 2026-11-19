"""
storage.py
- บันทึก/โหลดข้อมูลฟอร์มปัจจุบันเป็นไฟล์ JSON อัตโนมัติ (กันข้อมูลหาย)
- บันทึกประวัติทุกครั้งที่สั่งพิมพ์ ลงไฟล์ Excel (ฐานข้อมูล.xlsx)
"""

import json
import os
import sys
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
    from openpyxl import Workbook, load_workbook
except ImportError:
    openpyxl = None

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
AUTOSAVE_FILE = BASE_DIR / "ข้อมูลล่าสุด.json"
DATABASE_FILE = BASE_DIR / "ฐานข้อมูล.xlsx"

DB_HEADERS = [
    "ลำดับ", "วันที่-เวลา", "เอกสารที่พิมพ์", "ยศ", "ชื่อ-สกุล", "ปีเกิด",
    "เลขบัตรประชาชน", "บ้านเลขที่", "หมู่ที่", "ตำบล", "อำเภอ", "จังหวัด",
]


def autosave(data: dict) -> None:
    try:
        with open(AUTOSAVE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("บันทึกอัตโนมัติล้มเหลว:", e)


def load_autosave() -> dict:
    if AUTOSAVE_FILE.exists():
        try:
            with open(AUTOSAVE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print("โหลดข้อมูลอัตโนมัติล้มเหลว:", e)
    return {}


def log_to_database(doc_name: str, data: dict) -> None:
    """เพิ่มแถวใหม่ในไฟล์ ฐานข้อมูล.xlsx ทุกครั้งที่มีการสั่งพิมพ์เอกสาร"""
    if openpyxl is None:
        print("ไม่พบไลบรารี openpyxl - ข้ามการบันทึกฐานข้อมูล")
        return

    if DATABASE_FILE.exists():
        wb = load_workbook(DATABASE_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "บันทึกการพิมพ์"
        ws.append(DB_HEADERS)
        for cell in ws[1]:
            cell.font = openpyxl.styles.Font(bold=True)
        for col in "ABCDEFGHIJKL":
            ws.column_dimensions[col].width = 18

    existing_data_rows = max(ws.max_row - 1, 0)  # แถว 1 คือหัวตาราง
    seq = existing_data_rows + 1

    ws.append([
        seq,
        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        doc_name,
        data.get("yos", ""),
        data.get("name", ""),
        data.get("birth_year", ""),
        data.get("idcard", ""),
        data.get("address_no", ""),
        data.get("moo", ""),
        data.get("tambon", ""),
        data.get("amphoe", ""),
        data.get("province", ""),
    ])

    try:
        wb.save(DATABASE_FILE)
    except PermissionError:
        print("ไม่สามารถบันทึกฐานข้อมูลได้ (ไฟล์อาจเปิดค้างอยู่)")
