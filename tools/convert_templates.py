"""แปลงเทมเพลต .docx: เปลี่ยนข้อความหน่วยงาน/ผู้ลงนาม/วันที่ที่พิมพ์ตายตัว ให้เป็น <<placeholder>>
ใช้: python tools/convert_templates.py <โฟลเดอร์ต้นฉบับ> <โฟลเดอร์ผลลัพธ์>
(รันครั้งเดียวตอนเตรียมเทมเพลต ไม่ได้ใช้ตอนโปรแกรมทำงานปกติ)"""
import re
import shutil
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from docx_engine import _replace_in_xml_bytes, text_width  # noqa: E402

MON_ABBR = r"(?:ม\.ค\.|ก\.พ\.|มี\.ค\.|เม\.ย\.|พ\.ค\.|มิ\.ย\.|ก\.ค\.|ส\.ค\.|ก\.ย\.|ต\.ค\.|พ\.ย\.|ธ\.ค\.)"
MON_FULL = r"(?:มกราคม|กุมภาพันธ์|มีนาคม|เมษายน|พฤษภาคม|มิถุนายน|กรกฎาคม|สิงหาคม|กันยายน|ตุลาคม|พฤศจิกายน|ธันวาคม)"
D = r"[0-9๐-๙]"


def ph(key):
    return f"<<{key}>>"


def padded(key):
    """แทนทั้ง 'ช่องว่างนำหน้า+ข้อความ' ด้วย placeholder แบบรักษากึ่งกลาง"""
    def f(m):
        lead = len(m.group(1))
        core = m.group(0)[lead:]
        return f"<<{key}#{lead},{text_width(core)}>>"
    return f


def only(core_strip_equals, key):
    """แทนเฉพาะเมื่อทั้งย่อหน้ามีแค่ข้อความนี้ (กันชนกับที่อื่น)"""
    def f(m):
        return ph(key) if m.string.strip() == core_strip_equals else None
    return f


RULES = [
    # ---- ผู้ลงนาม (เต็ม) ----
    (r"(\s*)\(\s*วีรวัช\s+สุขใส\s*\)", padded("ชื่อผู้ลงนาม")),
    (r"(\s*)ผู้ช่วยสัสดีอำเภอวานรนิวาส\s*จังหวัดสกลนคร\s*รักษาราชการ", padded("ตำแหน่งผู้ลงนาม1")),
    (r"(\s*)สัสดีอำเภอวานรนิวาส\s*จังหวัดสกลนคร\s*ฝ่ายสรรพกำลัง", padded("ตำแหน่งผู้ลงนาม2")),
    (r"ร้อยโท", None),  # จัดการด้านล่างด้วย resolver เฉพาะ
    # ---- ผู้ลงนาม (ย่อ) / ผู้พิมพ์ ----
    (r"(\s*)\(\s*วีรวัชร์\s+สุขใส\s*\)", padded("ชื่อผู้ลงนามย่อ")),
    (r"(\s*)\(\s*ธนัญชัย\s+ทิลารักษ์\s*\)", padded("ชื่อผู้พิมพ์")),
    (r"(\s*)ผช\.สด\.อ\.วานรนิวาส\s*จว\.ส\.น\.\s*รักษาราชการ", padded("ตำแหน่งย่อ1")),
    (r"สด\.อ\.วานรนิวาส\s*จว\.ส\.น\.", only("สด.อ.วานรนิวาส จว.ส.น.", "ตำแหน่งย่อ2")),
    (r"ร\.ท\.", None),
    (r"ส\.อ\.\s*ธนัญชัย\s*ฯ(?=\s*พิมพ์/ทาน)", lambda m: ph("ผู้พิมพ์ย่อ")),
    (r"สิบเอก(?=\s*บันทึก/อ่าน)", lambda m: ph("ยศผู้พิมพ์")),
    # ---- ติดต่อ ----
    (r"โทร\.\s*0\s*5625\s*5802\s*ต่อ\s*31037", lambda m: "โทร. " + ph("โทรติดต่อ")),
    (r"sd\.mtc\.31@gmail\.com", lambda m: ph("อีเมล")),
    (r"โทร\.\s*[๐0]\s*4279\s*1834", lambda m: "โทร. " + ph("โทรหน่วย")),
    # ---- ที่อยู่/ชื่อหน่วย (ยาวก่อนสั้น) ----
    (r"หน่วยสัสดีอำเภอวานรนิวาส", lambda m: ph("ชื่อหน่วย")),
    (r"ศูนย์ราชการอำเภอ\s*สน\s*47120", lambda m: ph("ที่อยู่หน่วย")),
    (r"ที่ว่าการอำเภอวานรนิวาส", lambda m: ph("ที่ว่าการอำเภอ")),
    (r"ปลัดอำเภอวานรนิวาส", lambda m: ph("ปลัดอำเภอ")),
    (r"สถานีตำรวจภูธรวานรนิวาส", lambda m: ph("สถานีตำรวจ")),
    (r"ผู้ว่าราชการจังหวัดสกลนคร", lambda m: "ผู้ว่าราชการจังหวัด" + ph("จังหวัดหน่วย")),
    (r"อำเภอวานรนิวาส", lambda m: "อำเภอ" + ph("อำเภอหน่วย")),
    (r"หน่วย\s*สด\.อ\.วานรนิวาส\s*จว\.ส\.น\.", lambda m: ph("ส่วนราชการ")),
    (r"สด\.จว\.ส\.น\.", lambda m: "สด.จว." + ph("ตัวย่อจังหวัด")),
    (r"สน\s*0402/", lambda m: ph("ที่หนังสืออำเภอ")),
    (r"กห\s*๐๔๘๒\.๗๗\.๒๖\.3/", lambda m: ph("ที่หนังสือ")),
    # ---- วันที่อัตโนมัติ (ไม่มีวัน) ----
    (rf"/\s*25{D}{D}", lambda m: "/" + ph("ปีอารบิก")),
    (rf"({MON_FULL})(\s*พ\.ศ\.\s*){D}{{4}}", lambda m: ph("เดือนเต็ม") + re.sub(r"\S+", "", m.group(2), count=0) + "พ.ศ. " + ph("ปีเต็ม")),
    (rf"(?:{D}{{1,2}}\s*)?{MON_FULL}\s*{D}{{4}}", lambda m: ph("เดือนเต็มปีเต็ม")),
    (rf"(?:{D}{{1,2}}\s*)?{MON_ABBR}\s*{D}{{2}}(?!{D})", lambda m: ph("เดือนย่อปีย่อ")),
]


def _rank_resolver_full(m):
    return ph("ยศผู้ลงนาม") if m.string.strip() == "ร้อยโท" or re.search(r"ร้อยโท\s*(?=บันทึก/อ่าน)", m.string[m.start():]) else None


def _rank_resolver_short(m):
    s = m.string
    if s.strip() == "ร.ท." or re.match(r"ร\.ท\.\s*ร่าง/ตรวจ", s[m.start():]):
        return ph("ยศย่อ")
    return None


SDT_DATE = re.compile(r"<w:sdt>(?:(?!</w:sdt>).)*?<w:date\b.*?<w:sdtContent>(.*?)</w:sdtContent></w:sdt>", re.S)


def convert_part(xml: bytes) -> bytes:
    # ตัวควบคุมวันที่ของ Word (content control) ไม่อัปเดตเอง -> แกะออกเหลือข้อความธรรมดาก่อนแปลง
    xml = SDT_DATE.sub(lambda m: m.group(1), xml.decode("utf-8")).encode("utf-8")
    for pat, res in RULES:
        if res is None:
            res = _rank_resolver_full if pat == "ร้อยโท" else _rank_resolver_short
        xml = _replace_in_xml_bytes(xml, {}, pattern=re.compile(pat), resolver=res)
    return xml


def convert(src: Path, dst: Path):
    shutil.copyfile(src, dst)
    tmp = str(dst) + ".tmp"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == "word/document.xml" or re.match(r"word/(header|footer)\d*\.xml", it.filename):
                data = convert_part(data)
            zout.writestr(it, data)
    shutil.move(tmp, dst)


if __name__ == "__main__":
    src_dir, dst_dir = Path(sys.argv[1]), Path(sys.argv[2])
    dst_dir.mkdir(parents=True, exist_ok=True)
    for f in sorted(src_dir.glob("*.docx")):
        convert(f, dst_dir / f.name)
        print("converted", f.name)
