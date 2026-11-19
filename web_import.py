"""web_import.py - ดึงข้อมูลบุคคลจากหน้า "ตรวจสอบข้อมูลบุคคล" ของ TDC Recruit 4.0
(https://recruit4.tdc.mi.th/app/check-person-info) มาใส่ฟอร์ม SD8System

ส่วนที่ทดสอบได้โดยไม่ต้องมีเว็บ: parse_person()  (รับ dict ดิบจากหน้าเว็บ -> คืนค่าตามคีย์ของฟอร์ม)
ส่วนที่ต้องมีเว็บจริง: WebSession (Playwright)
"""
import os
import queue
import re
import threading

URL = "https://recruit4.tdc.mi.th/app/check-person-info"


def log_path():
    d = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "SD8System")
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        pass
    return os.path.join(d, "web_import_log.txt")


def dlog(msg):
    """บันทึกขั้นตอน/error ลงไฟล์ (ไม่เก็บข้อมูลบุคคล) ไว้ตามหาสาเหตุเวลาดึงไม่ได้"""
    try:
        import datetime
        import traceback
        with open(log_path(), "a", encoding="utf-8") as f:
            f.write(f"{datetime.datetime.now():%H:%M:%S} {msg}\n")
            if msg.startswith("ERROR"):
                f.write(traceback.format_exc() + "\n")
    except Exception:
        pass

# ต้องตั้งก่อน import playwright (เหมือน Auto Recruit) กัน Chromium หายตอนแพ็ค .exe
os.environ.setdefault(
    "PLAYWRIGHT_BROWSERS_PATH",
    os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "ms-playwright"),
)

# JS: อ่านทุก card -> {title, kv:{label:value}, subs:{ชื่อหัวข้อย่อย:{kv, muted}}}
SCRAPE_JS = r"""
() => {
  const t = (e) => (e ? e.innerText.replace(/\s+/g, " ").trim() : "");
  // หน้าเว็บอาจเรียงเป็น .kv หนึ่งก้อนที่มีคู่ .k/.v หลายคู่ต่อกัน (grid) หรือ .kv ต่อแถว -> รองรับทั้งสองแบบ
  const readKv = (root, skipSub) => {
    const kv = {};
    root.querySelectorAll(".kv").forEach((row) => {
      if (skipSub && row.closest(".subsection")) return;
      row.querySelectorAll(".k").forEach((kEl) => {
        let vEl = kEl.nextElementSibling;
        while (vEl && !vEl.classList.contains("v") && !vEl.classList.contains("k")) vEl = vEl.nextElementSibling;
        if (!vEl || !vEl.classList.contains("v")) return;
        const k = t(kEl), v = t(vEl);
        if (k && !(k in kv)) kv[k] = v;
      });
    });
    return kv;
  };
  const cards = [];
  document.querySelectorAll("app-check-person-info .grid > .card").forEach((c) => {
    const subs = {};
    c.querySelectorAll(".subsection").forEach((s) => {
      subs[t(s.querySelector(".subsection-title"))] = {
        kv: readKv(s, false), muted: t(s.querySelector(".muted")),
      };
    });
    cards.push({ title: t(c.querySelector(".section-title")), kv: readKv(c, true), subs });
  });
  return cards;
}
"""

PREFIXES = ["นางสาว", "น.ส.", "ด.ช.", "ด.ญ.", "นาย", "นาง", "เด็กชาย", "เด็กหญิง"]
THAI_MONTHS = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
               "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
THAI_MONTH_ABBR = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
                   "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]


def split_prefix(full: str):
    full = (full or "").strip()
    for p in PREFIXES:
        if full.startswith(p):
            return p, full[len(p):].strip()
    return "", full


def split_first_last(full: str):
    """ตัดคำนำหน้า แล้วแยกชื่อ/นามสกุลที่ช่องว่างแรก"""
    _, rest = split_prefix(full)
    parts = rest.split(None, 1)
    return (parts[0], parts[1]) if len(parts) == 2 else (rest, "")


def _after(word, text):
    """คืนข้อความหลังคำ word จนถึงคำบอกตำแหน่งถัดไป (หมู่ที่/ตำบล/อำเภอ/จังหวัด) หรือจบ"""
    m = re.search("(?:" + word + r")\s*(.+?)(?=\s+(?:หมู่ที่|หมู่|ตำบล|แขวง|อำเภอ|เขต|จังหวัด)\s|$)", text or "")
    return m.group(1).strip() if m else ""


def parse_address(kv):
    """kv ตัวอย่าง: บ้านเลขที่='192 หมู่ที่ 10', ถนน/ตำบล='ถนน - ตำบล เชิงชุม',
    อำเภอ/จังหวัด='อำเภอ พรรณานิคม จังหวัด สกลนคร'"""
    house = kv.get("บ้านเลขที่", "")
    m = re.match(r"\s*(\S+)(?:\s+หมู่(?:ที่)?\s*(\S+))?", house)
    no = m.group(1) if m else house
    moo = (m.group(2) or "") if m else ""
    if moo in ("-", "0"):
        moo = ""
    road_tambon = kv.get("ถนน/ตำบล", "")
    tambon = _after("ตำบล|แขวง", road_tambon)
    ap = kv.get("อำเภอ/จังหวัด", "")
    return {
        "address_no": no, "moo": moo, "tambon": tambon,
        "amphoe": _after("อำเภอ|เขต", ap), "province": _after("จังหวัด", ap),
    }


def parse_date(text):
    """รับ '12/03/2540' หรือ '12 มีนาคม 2540' หรือ '12 มี.ค. 2540' -> (วัน, เดือนไทยเต็ม, ปี พ.ศ.)
    ถ้าปีเป็น ค.ศ. (< 2400) จะบวก 543 ให้ ถ้าอ่านไม่ได้คืน None"""
    s = (text or "").strip()
    m = re.search(r"(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{4})", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.search(r"(\d{1,2})\s+(\S+)\s+(\d{4})", s)
        if not m:
            return None
        name = m.group(2)
        if name in THAI_MONTHS:
            mo = THAI_MONTHS.index(name) + 1
        elif name in THAI_MONTH_ABBR:
            mo = THAI_MONTH_ABBR.index(name) + 1
        else:
            return None
        d, y = int(m.group(1)), int(m.group(3))
    if not (1 <= mo <= 12 and 1 <= d <= 31):
        return None
    if y < 2400:
        y += 543
    return str(d), THAI_MONTHS[mo - 1], str(y)


def parse_person(cards):
    """cards = ผลจาก SCRAPE_JS -> (values, notes)
    values: dict คีย์เดียวกับช่องฟอร์ม (ใส่เฉพาะที่อ่านได้)
    notes: รายการข้อความเตือนสำหรับ log"""
    v, notes = {}, []
    def find(word):
        return next((c for c in cards if word in c["title"]), None)

    person = find("ข้อมูลบุคคล")
    if not person:
        return v, ["ไม่พบการ์ด 'ข้อมูลบุคคล' (ยังไม่ได้ค้นหา หรือไม่พบข้อมูล?)"]
    kv = person["kv"]

    idc = re.sub(r"\D", "", kv.get("เลข ปชช", ""))
    if idc:
        v["idcard"] = idc
    full = kv.get("ชื่อ", "")
    if full:
        _, rest = split_prefix(full)
        v["name"] = rest
    if kv.get("สัญชาติ"):
        v["nationality"] = kv["สัญชาติ"]

    # วันเกิด: ยังไม่ทราบป้ายแน่ชัด -> หาป้ายที่มีคำว่า 'เกิด' ก่อน ไม่งั้นลองทุกช่องที่ดูเป็นวันที่
    birth = None
    for k, val in kv.items():
        if "เกิด" in k:
            birth = parse_date(val)
            break
    if birth is None:
        for k, val in kv.items():
            if k in ("เลข ปชช",):
                continue
            birth = parse_date(val)
            if birth:
                notes.append(f"วันเกิดอ่านจากป้าย '{k}' (เดาจากรูปแบบวันที่)")
                break
    if birth:
        v["birth_day"], v["birth_month"], v["birth_yearfull"] = birth
        v["birth_year"] = birth[2]
    else:
        notes.append("อ่านวันเกิดไม่ได้")

    addr_card = find("ที่อยู่ปัจจุบัน")
    if addr_card:
        v.update({k: x for k, x in parse_address(addr_card["kv"]).items() if x})
        subs = addr_card["subs"]
    else:
        notes.append("ไม่พบการ์ด 'ที่อยู่ปัจจุบัน'")
        subs = {}
    for c in cards:  # บิดา/มารดา อาจอยู่คนละการ์ดกับที่อยู่
        subs = {**c["subs"], **subs}

    for title, key in (("บิดา", "father"), ("มารดา", "mother")):
        s = subs.get(title)
        if not s or not s["kv"].get("ชื่อ"):
            notes.append(f"{title}: {s['muted'] if s and s['muted'] else 'ไม่มีข้อมูล'}")
            continue
        first, last = split_first_last(s["kv"]["ชื่อ"])
        v[f"{key}_first"], v[f"{key}_last"] = first, last
    return v, notes


# ------------------------- เซสชันเบราว์เซอร์ (background thread) -------------------------

def ensure_chromium_installed(log=print):
    """ติดตั้ง Chromium ให้ Playwright อัตโนมัติ (ครั้งแรกต้องต่อเน็ต) — ใช้ได้ทั้งตอนรัน .py และ .exe"""
    import sys
    try:
        from playwright.__main__ import main as cli
    except Exception as e:
        raise RuntimeError(f"หาโมดูล playwright ไม่เจอ: {e}")
    old = sys.argv
    try:
        sys.argv = ["playwright", "install", "chromium"]
        log("กำลังตรวจสอบ/ติดตั้งเบราว์เซอร์ Chromium (ครั้งแรกอาจใช้เวลาสักครู่)...")
        try:
            cli()
        except SystemExit as e:
            if e.code not in (0, None):
                raise RuntimeError(f"ติดตั้ง Chromium ไม่สำเร็จ (exit code {e.code})")
    finally:
        sys.argv = old


class WebSession:
    """เปิดเบราว์เซอร์ค้างไว้ 1 ตัว ใช้ซ้ำได้หลายคน
    ใช้: open() -> (ผู้ใช้ Login เอง) -> fetch(เลขบัตร) -> (values, notes) หรือ raise"""

    def __init__(self, log=print):
        self.log = log
        self._cmd = queue.Queue()
        self._res = queue.Queue()
        self._thread = None
        self.ready = threading.Event()   # เบราว์เซอร์พร้อมแล้ว (หรือเปิดไม่สำเร็จ ดู self.error)
        self.error = None

    @property
    def is_open(self):
        return self._thread is not None and self._thread.is_alive()

    def open(self):
        if self.is_open:
            return
        self.ready.clear()
        self.error = None
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def fetch(self, citizen_id, timeout=75):
        """เรียกจากเธรดอื่น (ไม่ใช่ UI) — บล็อกจนได้ผล"""
        self._cmd.put(("fetch", citizen_id))
        try:
            kind, payload = self._res.get(timeout=timeout)
        except queue.Empty:
            raise RuntimeError(f"หมดเวลารอผลจากเว็บ ({timeout} วินาที) — เบราว์เซอร์อาจค้าง/ปิดไปแล้ว ลองกดปุ่มใหม่\n"
                               f"รายละเอียดอยู่ที่ {log_path()}")
        if kind == "err":
            raise RuntimeError(payload)
        return payload

    def close(self):
        self._cmd.put(("close", None))

    def _run(self):
        try:
            from playwright.sync_api import sync_playwright
            ensure_chromium_installed(self.log)
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=False)
                page = browser.new_context().new_page()
                page.goto(URL)
                dlog("run: เปิดเบราว์เซอร์สำเร็จ")
                self.ready.set()
                self.log("เปิดเว็บแล้ว — Login ให้เสร็จ แล้วกลับมากด 'ดึงจากเว็บ' อีกครั้ง")
                while browser.is_connected():
                    try:
                        cmd, arg = self._cmd.get(timeout=0.3)
                    except queue.Empty:
                        page.wait_for_timeout(100)
                        continue
                    if cmd == "close":
                        break
                    try:
                        self._res.put(("ok", self._fetch(page, arg)))
                    except Exception as e:
                        dlog(f"ERROR fetch: {e}")
                        self._res.put(("err", str(e)))
                try:
                    browser.close()
                except Exception:
                    pass
        except Exception as e:
            dlog(f"ERROR run: {e}")
            self.log(f"เปิดเบราว์เซอร์ไม่ได้: {e}")
            self.error = f"เปิดเบราว์เซอร์ไม่ได้: {e}\nรายละเอียดอยู่ที่ {log_path()}"
            self.ready.set()

    @staticmethod
    def _fetch(page, citizen_id):
        import re as _re
        dlog("fetch: start")
        if "check-person-info" not in page.url:
            dlog("fetch: goto (url ตอนนี้ไม่ใช่หน้าตรวจสอบ)")
            page.goto(URL)
        box = page.locator('input[placeholder="เลขประจำตัวประชาชน 13 หลัก"]')
        box.wait_for(state="visible", timeout=10000)
        box.fill("")
        box.fill(citizen_id)
        # ปุ่ม 'ค้นหา' เป๊ะ ๆ (ไม่ชนกับแท็บ 'ค้นหาด้วย...') ใช้ข้อความในปุ่ม ไม่พึ่งชื่อ accessible ที่อาจมีไอคอนปน
        btn = page.locator("button", has_text=_re.compile(r"ค้นหา\s*$"))
        try:
            btn.first.click(timeout=5000)
            dlog("fetch: คลิกปุ่มค้นหา")
        except Exception:
            dlog("fetch: คลิกปุ่มไม่ได้ -> กด Enter แทน")
            box.press("Enter")
        # รอจนข้อมูลในหน้าเป็นของเลขบัตรที่ค้น (กันอ่านผลของคนก่อนหน้า)
        import time
        end = time.time() + 20
        values, notes, last_idc = {}, [], ""
        while time.time() < end:
            page.wait_for_timeout(400)
            try:
                cards = page.evaluate(SCRAPE_JS)
            except Exception:
                continue
            if not cards:
                continue
            values, notes = parse_person(cards)
            last_idc = values.get("idcard", "")
            if last_idc == citizen_id:
                page.wait_for_timeout(400)  # ให้การ์ดบิดา/มารดาโหลดครบ
                cards = page.evaluate(SCRAPE_JS)
                values, notes = parse_person(cards)
                dlog(f"fetch: อ่านได้ {len(values)} ช่อง; หมายเหตุ={notes}")
                return values, notes
        if last_idc:
            raise RuntimeError(f"เลขบัตรในหน้าเว็บ ({last_idc}) ไม่ตรงกับที่ค้นหา ({citizen_id}) — ไม่นำมาใส่เพื่อกันผิดคน")
        raise RuntimeError("ไม่พบข้อมูลบุคคลหลังค้นหา (เลขบัตรผิด / ไม่มีข้อมูล / ยังไม่ได้ Login / อยู่ผิดหน้า)\n"
                           f"รายละเอียดอยู่ที่ {log_path()}")
