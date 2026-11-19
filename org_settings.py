"""org_settings.py - ตั้งค่าหน่วยงาน/ผู้ลงนาม (กรอกครั้งแรก แก้ภายหลังได้) + วันที่อัตโนมัติ
เก็บเป็นไฟล์ JSON ข้างโปรแกรม (ถ้าเขียนไม่ได้ใช้ LOCALAPPDATA) เขียนแบบ atomic มีเลขเวอร์ชัน
ไม่พึ่ง tkinter จึงทดสอบแยกได้"""

import json
import os
import sys
from datetime import datetime
from pathlib import Path

SETTINGS_VERSION = 1

if getattr(sys, "frozen", False):
    _BASE = Path(sys.executable).resolve().parent
else:
    _BASE = Path(__file__).resolve().parent
SETTINGS_NAME = "ตั้งค่าหน่วยงาน.json"

THAI_MONTHS = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
               "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
THAI_MONTH_ABBR = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
                   "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
_TH_DIGITS = str.maketrans("0123456789", "๐๑๒๓๔๕๖๗๘๙")

# (key, ป้ายกำกับ, ตัวอย่างที่แสดงเป็นคำใบ้, จำเป็นต้องกรอก)
GROUPS = [
    ("พื้นที่ (กรอกก่อน ช่องอื่นจะเติมให้อัตโนมัติ)", [
        ("amphoe", "อำเภอ", "วานรนิวาส", True),
        ("province", "จังหวัด", "สกลนคร", True),
        ("prov_abbr", "ตัวย่อจังหวัด", "ส.น.", True),
    ]),
    ("ที่อยู่เอกสาร / หน่วยงาน (หัวหนังสือรับรอง)", [
        ("unit_name", "ชื่อหน่วย", "หน่วยสัสดีอำเภอวานรนิวาส", True),
        ("unit_address", "ที่อยู่หน่วย", "ศูนย์ราชการอำเภอ สน 47120", True),
        ("unit_phone", "โทรศัพท์หน่วย", "๐ 4279 1834", True),
    ]),
    ("ส่วนราชการ / เลขที่หนังสือ", [
        ("agency", "ส่วนราชการ", "หน่วย สด.อ.วานรนิวาส จว.ส.น.", True),
        ("doc_no", "ที่หนังสือ", "กห ๐๔๘๒.๗๗.๒๖.3/", True),
        ("doc_no_amphoe", "ที่หนังสืออำเภอ", "สน 0402/", True),
        ("pk14_place", "สถานที่ ป.ค.14 (ที่ว่าการ)", "ที่ว่าการอำเภอวานรนิวาส", True),
        ("pk14_officer", "ปลัดอำเภอ (ป.ค.14)", "ปลัดอำเภอวานรนิวาส", True),
        ("police", "สถานีตำรวจ (ใน ป.ค.14)", "สถานีตำรวจภูธรวานรนิวาส", True),
    ]),
    ("ผู้ลงนามหนังสือ (ฉบับเต็ม — หนังสือรับรอง)", [
        ("sg_rank", "ยศ", "ร้อยโท", True),
        ("sg_name", "ชื่อ-นามสกุล", "วีรวัช  สุขใส", True),
        ("sg_pos1", "ตำแหน่งบรรทัด 1", "ผู้ช่วยสัสดีอำเภอวานรนิวาส จังหวัดสกลนคร รักษาราชการ", False),
        ("sg_pos2", "ตำแหน่งบรรทัด 2", "สัสดีอำเภอวานรนิวาส จังหวัดสกลนคร ฝ่ายสรรพกำลัง", False),
        ("sg_tel", "โทร (ท้ายหนังสือรับรอง)", "0 5625 5802 ต่อ 31037", False),
        ("sg_mail", "อีเมล (ท้ายหนังสือรับรอง)", "sd.mtc.31@gmail.com", False),
    ]),
    ("ผู้ลงนามย่อ (บันทึกข้อความ สด.43 / ป.ค.14)", [
        ("ss_rank", "ยศย่อ", "ร.ท.", True),
        ("ss_name", "ชื่อ-นามสกุล", "วีรวัชร์  สุขใส", True),
        ("ss_pos1", "ตำแหน่งย่อ บรรทัด 1", "ผช.สด.อ.วานรนิวาส จว.ส.น. รักษาราชการ", False),
        ("ss_pos2", "ตำแหน่งย่อ บรรทัด 2", "สด.อ.วานรนิวาส จว.ส.น.", False),
    ]),
    ("เจ้าหน้าที่พิมพ์ (ท้ายหนังสือ / ป.ค.14)", [
        ("tp_rank", "ยศ", "สิบเอก", False),
        ("tp_name", "ชื่อ-นามสกุล", "ธนัญชัย  ทิลารักษ์", False),
        ("tp_short", "ชื่อย่อ (บรรทัดพิมพ์/ทาน)", "ส.อ. ธนัญชัย ฯ", False),
    ]),
]
ALL_KEYS = [k for _, items in GROUPS for k, _, _, _ in items]
REQUIRED_KEYS = [k for _, items in GROUPS for k, _, _, req in items if req]
LABELS = {k: label for _, items in GROUPS for k, label, _, _ in items}
HINTS = {k: hint for _, items in GROUPS for k, _, hint, _ in items}

# ค่าตัวอย่าง (ปุ่ม "โหลดค่าตัวอย่าง") = ค่าที่เทมเพลตเดิมใช้อยู่
SAMPLE = dict(HINTS)


def derive(s: dict) -> dict:
    """ค่าที่ประกอบขึ้นได้จาก อำเภอ/จังหวัด/ตัวย่อจังหวัด (ใช้เติมช่องที่ยังว่าง)"""
    a, p, ab = s.get("amphoe", "").strip(), s.get("province", "").strip(), s.get("prov_abbr", "").strip()
    if not a:
        return {}
    d = {
        "unit_name": f"หน่วยสัสดีอำเภอ{a}",
        "pk14_place": f"ที่ว่าการอำเภอ{a}",
        "pk14_officer": f"ปลัดอำเภอ{a}",
        "police": f"สถานีตำรวจภูธร{a}",
        "ss_pos2": f"สด.อ.{a} จว.{ab}",
        "ss_pos1": f"ผช.สด.อ.{a} จว.{ab} รักษาราชการ",
        "agency": f"หน่วย สด.อ.{a} จว.{ab}",
    }
    if p:
        d["sg_pos1"] = f"ผู้ช่วยสัสดีอำเภอ{a} จังหวัด{p} รักษาราชการ"
        d["sg_pos2"] = f"สัสดีอำเภอ{a} จังหวัด{p} ฝ่ายสรรพกำลัง"
    return d


def settings_path(for_write: bool = False) -> Path:
    p = _BASE / SETTINGS_NAME
    if p.exists() or not for_write:
        return p
    try:
        _BASE.mkdir(parents=True, exist_ok=True)
        probe = _BASE / ".write_test"
        probe.write_text("x")
        probe.unlink()
        return p
    except Exception:
        alt = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "SD8System"
        alt.mkdir(parents=True, exist_ok=True)
        return alt / SETTINGS_NAME


def _find_existing() -> Path | None:
    for p in (_BASE / SETTINGS_NAME,
              Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "SD8System" / SETTINGS_NAME):
        if p.exists():
            return p
    return None


def load() -> dict:
    """คืน dict ครบทุกคีย์ (คีย์ที่ไม่มี = ''), ไฟล์เสีย/ไม่มี = ค่าว่างทั้งหมด"""
    data = {}
    p = _find_existing()
    if p:
        try:
            raw = json.loads(p.read_text(encoding="utf-8"))
            data = raw.get("values", {}) if isinstance(raw, dict) else {}
        except Exception:
            data = {}
    return {k: str(data.get(k, "") or "") for k in ALL_KEYS}


def save(values: dict) -> Path:
    path = settings_path(for_write=True)
    payload = {"version": SETTINGS_VERSION,
               "saved": datetime.now().isoformat(timespec="seconds"),
               "values": {k: str(values.get(k, "") or "").strip() for k in ALL_KEYS}}
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)
    return path


def missing_required(values: dict) -> list:
    return [LABELS[k] for k in REQUIRED_KEYS if not str(values.get(k, "")).strip()]


def is_configured() -> bool:
    return _find_existing() is not None and not missing_required(load())


def to_thai_digits(s) -> str:
    return str(s).translate(_TH_DIGITS)


def date_mapping(now: datetime | None = None) -> dict:
    """วันที่อัตโนมัติ (ไม่มี 'วัน' ผู้ใช้เขียนเองในช่องว่าง) ใช้ปี พ.ศ."""
    now = now or datetime.now()
    be = now.year + 543
    m = now.month - 1
    return {
        "เดือนเต็ม": THAI_MONTHS[m],
        "ปีเต็ม": to_thai_digits(be),
        "เดือนเต็มปีเต็ม": f"{THAI_MONTHS[m]} {to_thai_digits(be)}",
        "เดือนย่อปีย่อ": f"{THAI_MONTH_ABBR[m]} {to_thai_digits(be % 100):0>2}",
        "ปีอารบิก": str(be),
    }


def org_mapping(values: dict, now: datetime | None = None) -> dict:
    """คีย์ <<placeholder>> ของหน่วยงาน/ผู้ลงนาม/วันที่ ที่เทมเพลตใช้"""
    v = {k: str(values.get(k, "") or "").strip() for k in ALL_KEYS}
    m = {
        "ชื่อหน่วย": v["unit_name"],
        "ที่อยู่หน่วย": v["unit_address"],
        "โทรหน่วย": v["unit_phone"],
        "ส่วนราชการ": v["agency"],
        "ที่หนังสือ": v["doc_no"],
        "ที่หนังสืออำเภอ": v["doc_no_amphoe"],
        "อำเภอหน่วย": v["amphoe"],
        "จังหวัดหน่วย": v["province"],
        "ตัวย่อจังหวัด": v["prov_abbr"],
        "ที่ว่าการอำเภอ": v["pk14_place"],
        "ปลัดอำเภอ": v["pk14_officer"],
        "สถานีตำรวจ": v["police"],
        "ยศผู้ลงนาม": v["sg_rank"],
        "ชื่อผู้ลงนาม": f"( {v['sg_name']} )" if v["sg_name"] else "",
        "ตำแหน่งผู้ลงนาม1": v["sg_pos1"],
        "ตำแหน่งผู้ลงนาม2": v["sg_pos2"],
        "ตำแหน่งเต็ม1": v["sg_pos1"],
        "ตำแหน่งเต็ม2": v["sg_pos2"],
        "โทรติดต่อ": v["sg_tel"],
        "อีเมล": v["sg_mail"],
        "ยศย่อ": v["ss_rank"],
        "ชื่อผู้ลงนามย่อ": f"( {v['ss_name']} )" if v["ss_name"] else "",
        "ตำแหน่งย่อ1": v["ss_pos1"],
        "ตำแหน่งย่อ2": v["ss_pos2"],
        "ยศผู้พิมพ์": v["tp_rank"],
        "ชื่อผู้พิมพ์": f"( {v['tp_name']} )" if v["tp_name"] else "",
        "ผู้พิมพ์ย่อ": v["tp_short"],
    }
    m.update(date_mapping(now))
    return m
