"""
docx_engine.py
เครื่องมือเติมข้อมูลลงไฟล์ Word (.docx) โดยแทนที่ placeholder รูปแบบ <<ชื่อฟิลด์>>
ทำงานได้แม้ Word จะแยกข้อความ <<ชื่อ>> ออกเป็นหลาย run (พบบ่อยเวลาพิมพ์ในกล่องข้อความ)
ไม่ต้องพึ่ง Microsoft Word หรือ COM Automation - ทำงานได้ทั้งบน Windows/Mac/Linux
"""

import re
import shutil
import unicodedata
import zipfile
from pathlib import Path
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}

PLACEHOLDER_RE = re.compile(r"<<[^<>]+>>")

# อัตราส่วนความกว้างตัวอักษรไทย เทียบกับ "ช่องว่าง" 1 ตัว (ใช้ชดเชยการจัดกึ่งกลางด้วยช่องว่างในเทมเพลต)
SPACE_RATIO = 1.5


def text_width(s: str) -> int:
    """นับจำนวนตัวอักษรที่กินที่จริง (ไม่นับสระ/วรรณยุกต์ที่ลอยอยู่บน-ล่าง)"""
    return sum(1 for ch in s if unicodedata.category(ch) != "Mn")


def _padded(value: str, spec: str) -> str:
    """คีย์รูปแบบ <<ชื่อ#lead,w>> : lead = จำนวนช่องว่างนำหน้าเดิมในเทมเพลต, w = ความกว้างข้อความเดิม
    คืนค่า = ช่องว่างนำหน้า (ปรับตามส่วนต่างความกว้างเพื่อให้ยังอยู่กึ่งกลางเดิม) + value"""
    try:
        lead, w = (int(x) for x in spec.split(","))
    except ValueError:
        return value
    pad = max(0, lead + round((w - text_width(value)) * SPACE_RATIO / 2))
    return " " * pad + value


def _iter_target_parts(zin: zipfile.ZipFile):
    """คืนรายชื่อไฟล์ XML ภายใน .docx ที่ควรค้นหา/แทนที่ข้อความ
    (เนื้อหาหลัก + กล่องข้อความที่ฝังอยู่ใน document.xml + header/footer)"""
    names = zin.namelist()
    targets = []
    for n in names:
        if n == "word/document.xml":
            targets.append(n)
        elif re.match(r"word/(header|footer)\d*\.xml", n):
            targets.append(n)
    return targets


def _direct_runs(p):
    """คืนเฉพาะ run (w:r) ที่เป็นของย่อหน้านี้จริง ๆ เท่านั้น
    (ไม่ไล่ลงไปในกล่องข้อความ/w:drawing ที่แขวนอยู่ในย่อหน้า เพราะกล่องข้อความ
    มีย่อหน้า/รันเป็นของตัวเอง ถูกประมวลผลแยกต่างหากอยู่แล้วในลูปหลัก
    ถ้าไม่กันจุดนี้ไว้ จะทำให้ข้อความถูกแทนที่ซ้ำซ้อนและเพี้ยน)"""
    runs = []
    for child in p:
        tag = etree.QName(child).localname
        if tag == "r":
            runs.append(child)
        elif tag in ("hyperlink", "ins", "del", "smartTag"):
            # กรณี run ถูกห่อด้วยแท็กเหล่านี้ (พบได้บ่อย) ให้ดึง run ชั้นถัดไปมาด้วย
            for gc in child:
                if etree.QName(gc).localname == "r":
                    runs.append(gc)
    return runs


def _replace_in_xml_bytes(xml_bytes: bytes, mapping: dict, pattern=PLACEHOLDER_RE, resolver=None, skip_time_paras=False) -> bytes:
    """pattern/resolver ใช้ตอนแปลงเทมเพลต (หาข้อความตรง ๆ แล้วแทนด้วย placeholder);
    ตอนเติมข้อมูลปกติใช้ค่าเริ่มต้น"""
    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(xml_bytes, parser=parser)

    # ค้นหาทุก <w:p> รวมถึงที่ซ้อนอยู่ในกล่องข้อความ (w:txbxContent)
    # แต่ละย่อหน้าจะถูกประมวลผล "เฉพาะ run ของตัวเอง" เท่านั้น (ดู _direct_runs)
    paragraphs = root.findall(".//w:p", NSMAP)

    for p in paragraphs:
        runs = _direct_runs(p)
        if not runs:
            continue
        if skip_time_paras and any("TIME" in (i.text or "") for i in p.iter("{%s}instrText" % NSMAP["w"])):
            continue  # ย่อหน้าที่มีฟิลด์วันที่ของ Word: ห้ามแตะ

        # เก็บข้อความของแต่ละ run (รองรับ run ที่ไม่มี w:t ด้วย เช่น w:br)
        run_texts = []
        for r in runs:
            t_el = r.find("w:t", NSMAP)
            run_texts.append(t_el.text if (t_el is not None and t_el.text) else "")

        full_text = "".join(run_texts)
        if pattern is PLACEHOLDER_RE and "<<" not in full_text:
            continue

        matches = list(pattern.finditer(full_text))
        if not matches:
            continue

        # แผนที่ตำแหน่งตัวอักษร -> (run index, offset ภายใน run)
        # เพื่อหาว่า placeholder แต่ละตัวครอบคลุม run ใดบ้าง
        run_bounds = []
        pos = 0
        for txt in run_texts:
            run_bounds.append((pos, pos + len(txt)))
            pos += len(txt)

        # ประมวลผลจากท้ายไปหน้า เพื่อไม่ให้ index ของ match อื่นเพี้ยนหลังแก้ไข
        for m in reversed(matches):
            if resolver is not None:
                replacement = resolver(m)
                if replacement is None:
                    continue
            else:
                key = m.group(0)[2:-2]  # ตัด << >> ออก
                if "#" in key:
                    base, spec = key.split("#", 1)
                    if base not in mapping:
                        continue
                    replacement = _padded(mapping[base], spec)
                else:
                    if key not in mapping:
                        continue  # ไม่รู้จักคีย์นี้ ข้ามไป (คงข้อความเดิมไว้)
                    replacement = mapping[key]

            start, end = m.start(), m.end()

            # หา run แรกและ run สุดท้ายที่ placeholder ครอบคลุมถึง
            first_idx = None
            last_idx = None
            for i, (rs, re_) in enumerate(run_bounds):
                if rs <= start < re_ or (rs == re_ == start):
                    first_idx = i if first_idx is None else first_idx
                if rs < end <= re_:
                    last_idx = i
            if first_idx is None:
                # placeholder เริ่มที่ตำแหน่งพอดีปลาย run ก่อนหน้า (edge case) -> หา run ถัดไป
                for i, (rs, re_) in enumerate(run_bounds):
                    if rs >= start:
                        first_idx = i
                        break
            if last_idx is None:
                last_idx = first_idx

            first_run = runs[first_idx]
            first_t = first_run.find("w:t", NSMAP)
            prefix = run_texts[first_idx][: start - run_bounds[first_idx][0]]

            last_run = runs[last_idx]
            last_t = last_run.find("w:t", NSMAP)
            suffix = run_texts[last_idx][end - run_bounds[last_idx][0]:]

            # เขียนค่าใหม่ลง run แรก (prefix + ค่าที่แทนที่ + suffix ถ้าอยู่ run เดียวกัน)
            if first_idx == last_idx:
                new_text = prefix + replacement + suffix
                if first_t is None:
                    first_t = etree.SubElement(first_run, f"{{{W_NS}}}t")
                first_t.text = new_text
                first_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            else:
                if first_t is None:
                    first_t = etree.SubElement(first_run, f"{{{W_NS}}}t")
                first_t.text = prefix + replacement
                first_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

                if last_t is None:
                    last_t = etree.SubElement(last_run, f"{{{W_NS}}}t")
                last_t.text = suffix
                last_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

                # ล้างข้อความของ run ที่อยู่ระหว่างกลาง (ถ้ามี)
                for mid_idx in range(first_idx + 1, last_idx):
                    mid_t = runs[mid_idx].find("w:t", NSMAP)
                    if mid_t is not None:
                        mid_t.text = ""

            # อัปเดต run_texts / run_bounds ในหน่วยความจำ ให้สอดคล้องกับการแก้ไขนี้
            # (จำเป็นสำหรับ match ก่อนหน้าที่ยังไม่ประมวลผลในลูป reversed เดียวกัน)
            if first_idx == last_idx:
                run_texts[first_idx] = prefix + replacement + suffix
            else:
                run_texts[first_idx] = prefix + replacement
                run_texts[last_idx] = suffix
                for mid_idx in range(first_idx + 1, last_idx):
                    run_texts[mid_idx] = ""
            pos = 0
            new_bounds = []
            for txt in run_texts:
                new_bounds.append((pos, pos + len(txt)))
                pos += len(txt)
            run_bounds = new_bounds

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


_TH_DIG = "๐๑๒๓๔๕๖๗๘๙"
_MON_FULL = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
_MON_ABBR = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."]


def _fmt_time_field(fmt: str, now, thai_digits: bool) -> str:
    be = now.year + 543
    out, i = [], 0
    while i < len(fmt):
        for tok, val in (("ดดดด", _MON_FULL[now.month-1]), ("ดดด", _MON_ABBR[now.month-1]),
                         ("ปปปป", str(be)), ("ปป", str(be)[-2:]), ("bb", str(be)[-2:]),
                         ("dd", "%02d" % now.day), ("ว", str(now.day))):
            if fmt.startswith(tok, i):
                out.append(val); i += len(tok); break
        else:
            out.append(fmt[i]); i += 1
    t = "".join(out)
    if thai_digits:
        t = "".join(_TH_DIG[int(c)] if c.isdigit() else c for c in t)
    return t


def _refresh_time_fields(xml_bytes: bytes) -> bytes:
    """เขียนวันที่ปัจจุบันลงผลลัพธ์ที่แคชของฟิลด์ TIME (Word จะอัปเดตเองเมื่อเปิดอยู่แล้ว
    แต่ตัวดูตัวอย่าง/โปรแกรมอื่นจะเห็นค่าที่แคชไว้)"""
    import datetime
    if b"TIME" not in xml_bytes:
        return xml_bytes
    now = datetime.datetime.now()
    W = "{%s}" % NSMAP["w"]
    root = etree.fromstring(xml_bytes, etree.XMLParser(remove_blank_text=False))
    for p in root.iter(W + "p"):
        state, instr, res = 0, "", []
        for r in p.iter(W + "r"):
            fc = r.find("w:fldChar", NSMAP)
            if fc is not None:
                t = fc.get(W + "fldCharType")
                if t == "begin":
                    state, instr, res = 1, "", []
                elif t == "separate" and state == 1:
                    state = 2
                elif t == "end" and state:
                    m = re.search(r'TIME\s*\\@\s*"([^"]*)"', instr)
                    ts = [x for x in res if x.text and x.text.strip()]
                    if m and ts:
                        th = any(c in _TH_DIG for x in ts for c in x.text)
                        ts[0].text = _fmt_time_field(m.group(1), now, th)
                        for x in ts[1:]:
                            x.text = ""
                    state = 0
                continue
            if state == 1:
                for it in r.findall("w:instrText", NSMAP):
                    instr += it.text or ""
            elif state == 2:
                res.extend(r.findall("w:t", NSMAP))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def fill_docx(template_path: str, output_path: str, mapping: dict) -> None:
    """
    เปิดไฟล์ template_path (.docx) แทนที่ <<key>> ทุกจุดด้วยค่าจาก mapping (dict: key -> str)
    บันทึกผลลัพธ์เป็น output_path
    ครอบคลุมเนื้อหาหลัก, กล่องข้อความ, header, footer
    """
    template_path = str(template_path)
    output_path = str(output_path)

    # คัดลอกไฟล์ต้นฉบับไปเป็นไฟล์ผลลัพธ์ก่อน แล้วแก้ไข part ที่จำเป็นภายใน zip
    shutil.copyfile(template_path, output_path)

    with zipfile.ZipFile(template_path, "r") as zin:
        targets = _iter_target_parts(zin)
        replacements = {}
        for name in targets:
            xml_bytes = zin.read(name)
            new_xml = _refresh_time_fields(_replace_in_xml_bytes(xml_bytes, mapping))
            replacements[name] = new_xml

    # เขียนไฟล์ zip ใหม่ทั้งหมด แทนที่เฉพาะ part ที่แก้ไข ส่วนที่เหลือคัดลอกเหมือนเดิม
    tmp_output = output_path + ".tmp"
    with zipfile.ZipFile(template_path, "r") as zin, \
         zipfile.ZipFile(tmp_output, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replacements.get(item.filename, zin.read(item.filename))
            zout.writestr(item, data)

    shutil.move(tmp_output, output_path)
