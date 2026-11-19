"""utils.py - ฟังก์ชันช่วยจัดรูปแบบข้อมูลและสร้าง mapping สำหรับแต่ละเอกสาร"""

import re


def format_thai_id(raw: str) -> str:
    """แปลงเลขบัตรประชาชน 13 หลัก เป็นรูปแบบ 0 0000 00000 00 0"""
    digits = re.sub(r"\D", "", raw or "")
    if len(digits) != 13:
        return raw or ""
    return f"{digits[0]} {digits[1:5]} {digits[5:10]} {digits[10:12]} {digits[12]}"


def combine_name(prefix: str, first: str, last: str) -> str:
    first = (first or "").strip()
    last = (last or "").strip()
    return f"{prefix}{first}  {last}"


def compute_age(birth_year: str, ref_year: str) -> str:
    try:
        return str(int(ref_year) - int(birth_year))
    except (ValueError, TypeError):
        return ""


def build_mapping(data: dict) -> dict:
    """สร้าง mapping รวมทุกคีย์ placeholder ที่เป็นไปได้ จากข้อมูลฟอร์ม (dict)
    ใช้ร่วมกันได้ทั้ง 4 เอกสาร - เอกสารไหนไม่ใช้คีย์ไหนก็แค่ไม่ถูกอ้างอิงในเทมเพลต"""

    father_name = combine_name("นาย", data.get("father_first", ""), data.get("father_last", ""))
    mother_name = combine_name("นาง", data.get("mother_first", ""), data.get("mother_last", ""))

    tb_combine = f"{data.get('tb_prefix', '')}{data.get('tb_year', '')}"
    jw_combine = f"{data.get('jw_prefix', '')}{data.get('jw_number', '')}"

    id_formatted = format_thai_id(data.get("idcard", ""))

    age = compute_age(data.get("birth_year", ""), data.get("ref_year", ""))

    birth_date_full = " ".join([
        (data.get("birth_day") or "").strip(),
        (data.get("birth_month") or "").strip(),
        (data.get("birth_yearfull") or "").strip(),
    ]).strip()

    release_date_full = " ".join([
        (data.get("release_day") or "").strip(),
        (data.get("release_month") or "").strip(),
        (data.get("release_year") or "").strip(),
    ]).strip()

    return {
        "ยศ": data.get("yos", ""),
        "ชื่อ": data.get("name", ""),
        "ปีเกิด": data.get("birth_year", ""),
        "อายุ": age,
        "ชื่อพ่อ": father_name,
        "ชื่อแม่": mother_name,
        "บ้านเลขที่": data.get("address_no", ""),
        "หมู่ที่": data.get("moo", ""),
        "ตำบล": data.get("tambon", ""),
        "อำเภอ": data.get("amphoe", ""),
        "จังหวัด": data.get("province", ""),
        "เลขบัตร": id_formatted,
        "ทบ": tb_combine,
        "จว": jw_combine,
        "วันเดือนปีเกิด": birth_date_full,
        "ผลัด": data.get("conscript_batch", ""),
        "สังกัด": data.get("unit", ""),
        "ประจำปี": data.get("conscript_year", ""),
        "วันปลด": release_date_full,
        # ---- ชุดข้อมูลผลตรวจเลือกทหาร (ใช้กับ CFRT / NEWRT) ----
        "วันที่เกณฑ์": data.get("draft_date", ""),
        "บัญชี": data.get("roll_no", ""),
        "จำนวนนับที่": data.get("count_no", ""),
        "เลขที่ 43": data.get("sd43_no", ""),
        "คนจำพวก": data.get("class_type", ""),
        "ส่วนสูง": data.get("height_cm", ""),
        "รอบตัว": data.get("chest_cm", ""),
        "จับสลากถูก": data.get("lottery_result", ""),
    }


# รายชื่อฟิลด์ทั้งหมด (key, ป้ายกำกับ, ค่าเริ่มต้น) - ใช้สำหรับ autosave / clear form
ALL_FIELD_KEYS = [
    "yos", "name", "birth_year", "ref_year",
    "nationality", "ethnicity",
    "address_no", "moo", "tambon", "amphoe", "province",
    "father_first", "father_last", "mother_first", "mother_last",
    "idcard",
    "tb_prefix", "tb_year", "jw_prefix", "jw_number",
    "birth_day", "birth_month", "birth_yearfull",
    "unit", "conscript_batch", "conscript_year",
    "release_day", "release_month", "release_year",
    "draft_date", "roll_no", "count_no", "sd43_no",
    "class_type", "height_cm", "chest_cm", "lottery_result",
]
