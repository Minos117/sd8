"""
app.py - ระบบขอ สด.8 ใหม่ / ป.ค.14 / หนังสือรับรอง (Desktop App)
UI ทันสมัยด้วย CustomTkinter, บันทึกข้อมูลอัตโนมัติกันข้อมูลหาย
เติมข้อมูลลงไฟล์ Word (.docx) โดยตรง ไม่ต้องพึ่ง Microsoft Word เปิดอยู่เบื้องหลัง
"""

import os
import sys
import platform
import subprocess
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import messagebox
import threading
import update_check
import linkage_check
import tkinter.font as tkfont
from PIL import Image

# ---------------------------------------------------------------
# แก้ปัญหาตัวอักษรเบลอ/หยักบน Windows (DPI awareness)
# ต้องเรียกก่อนสร้างหน้าต่างใด ๆ ทั้งหมด
# ---------------------------------------------------------------
if platform.system() == "Windows":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

# ---------------------------------------------------------------
# ค่าคงที่ / พาธไฟล์ (รองรับทั้งตอนรันเป็น .py และตอนถูกแพ็คเป็น .exe)
# ---------------------------------------------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent


def _write_crash_log(context: str):
    """บันทึกรายละเอียด error ลงไฟล์ crash_log.txt ข้าง exe/app.py

    สำคัญมากตอนแพ็คเป็น .exe แบบ --windowed เพราะโปรแกรมจะไม่มีหน้าต่าง
    console ให้เห็น error เลย — ถ้าไม่ log ไว้ โปรแกรมจะแค่ "หายไปเฉยๆ"
    โดยไม่บอกอะไรผู้ใช้เลย
    """
    try:
        log_path = BASE_DIR / "crash_log.txt"
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"\n{'=' * 60}\n{datetime.now()}  [{context}]\n")
            f.write(traceback.format_exc())
        return log_path
    except Exception:
        return None


# ---------------------------------------------------------------
# โหลดไลบรารีหลัก (customtkinter + โมดูลภายในโปรแกรม) อย่างปลอดภัย
# ถ้าโหลดไม่สำเร็จ (เช่น ไฟล์ .exe แพ็คไม่ครบ, ขาด hidden-import)
# ต้องแจ้งเตือนผู้ใช้ด้วย messagebox แทนการปล่อยให้โปรแกรมปิดตัวเงียบๆ
# ("UI หาย" ส่วนใหญ่เกิดจากจุดนี้)
# ---------------------------------------------------------------
try:
    import customtkinter as ctk

    from docx_engine import fill_docx
    from utils import build_mapping, ALL_FIELD_KEYS
    import storage
    import org_settings
    from settings_dialog import SettingsDialog, IdPromptDialog
except Exception:
    _log_path = _write_crash_log("STARTUP IMPORT ERROR")
    _root = None
    try:
        import tkinter as _tk
        _root = _tk.Tk()
        _root.withdraw()
        messagebox.showerror(
            "เปิดโปรแกรมไม่สำเร็จ",
            "โปรแกรมไม่สามารถโหลดไลบรารีที่จำเป็นได้ จึงไม่แสดงหน้าจอ\n\n"
            + traceback.format_exc(limit=3)
            + (f"\n\nรายละเอียดทั้งหมดถูกบันทึกไว้ที่:\n{_log_path}" if _log_path else ""),
        )
    except Exception:
        pass
    finally:
        if _root is not None:
            try:
                _root.destroy()
            except Exception:
                pass
    sys.exit(1)

def _load_bundled_fonts():
    """โหลดฟอนต์ใน fonts/ (Chakra Petch รองรับไทย, Press Start 2P สไตล์ 8-bit) ให้ใช้ได้เฉพาะโปรแกรมนี้
    ถ้าโหลดไม่ได้ (ไม่ใช่ Windows / ไม่มีไฟล์) ก็ข้ามไป แล้วใช้ฟอนต์สำรองตามลำดับเดิม"""
    if platform.system() != "Windows":
        return
    try:
        import ctypes
        for f in sorted((BASE_DIR / "fonts").glob("*.ttf")):
            ctypes.windll.gdi32.AddFontResourceExW(str(f), 0x10, 0)  # FR_PRIVATE
    except Exception:
        pass


_load_bundled_fonts()

APP_VER = "0.9.6"  # Beta — ต้องตรงกับ installer.iss (AppVersion) และ update.json
TEMPLATE_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# โลโก้ AMGF1: วางไฟล์ amgf1_logo.png ไว้ข้าง app.py
LOGO_PATH = BASE_DIR / "sd8_logo.png"
if not LOGO_PATH.exists():
    LOGO_PATH = BASE_DIR / "amgf1_logo.png"

# AMGF1 Theme — Mercedes-AMG Petronas F1 (dark / petronas teal / silver)
# ธีมมืดสนิท ตัวอักษรสว่างตัดกับพื้นหลังดำ ให้ความรู้สึกพรีเมียม/สปอร์ต
COLOR_BG = "#0B0B1A"
COLOR_SURFACE = "#14142B"
COLOR_PANEL_HEADER = "#0F0F22"
COLOR_CARD = "#14142B"
COLOR_CARD_ALT = "#1B1B38"
COLOR_BORDER = "#3A3A8C"
COLOR_LABEL_CHIP = "#2B2B6B"
COLOR_LABEL_TEXT = "#EDEDFF"
COLOR_ENTRY_BG = "#0B0B1A"
COLOR_ENTRY_BORDER = "#5B5BC8"
COLOR_ENTRY_TEXT = "#FFFFFF"
COLOR_GOLD = "#FFD84A"        # เหลืองอาร์เคด (สีเน้นหลัก)
COLOR_GOLD_HOVER = "#FFE98A"  # เหลืองสว่างขึ้นตอน hover
COLOR_SILVER = "#B8B8E0"
COLOR_BTN_PRIMARY = "#FFD84A"
COLOR_BTN_SECONDARY = "#22224A"
COLOR_BTN_TERTIARY = "#2E2E66"
COLOR_TEXT_LIGHT = "#EDEDFF"
COLOR_TEXT_MUTED = "#8F8FC0"
COLOR_TAB_SELECTED = "#5B5BFF"   # แท็บที่เลือก (ตัวอักษรขาวอ่านชัด)
COLOR_TAB_HOVER = "#7B7BFF"

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ---------------------------------------------------------------
# ฟอนต์: เลือกฟอนต์ที่มีอยู่จริงในเครื่องโดยอัตโนมัติ เรียงลำดับความสวย/คมชัด
# (ป้องกันปัญหาตัวอักษรหยัก/ตาลาย เวลาฟอนต์ที่ระบุไว้ไม่มีในเครื่อง)
# ---------------------------------------------------------------
UI_FONT_FAMILY = "TkDefaultFont"
_FONT_PRIORITY = ["Chakra Petch", "Leelawadee UI", "TH Sarabun New", "Tahoma", "Segoe UI", "Noto Sans Thai"]


def _detect_font(root) -> str:
    try:
        available = set(tkfont.families(root))
    except Exception:
        return "TkDefaultFont"
    for f in _FONT_PRIORITY:
        if f in available:
            return f
    return "TkDefaultFont"


def pixel_font(size=14):
    """ฟอนต์พิกเซล 8-bit (เฉพาะอักษรอังกฤษ/ตัวเลข ห้ามใช้กับข้อความไทย)"""
    return ctk.CTkFont(family="Press Start 2P", size=size)


def ui_font(size=16, weight="normal"):
    return ctk.CTkFont(family=UI_FONT_FAMILY, size=size, weight=weight)


DOCS = {
    "PK14": {"file": "PK14.docx", "label": "ป.ค.14", "group": "ขอใบแทน สด.8"},
    "SD8": {"file": "SD8.docx", "label": "ปะหน้า (สด.8)", "group": "ขอใบแทน สด.8"},
    "CFSD": {"file": "CFSD.docx", "label": "รับรองพลทหาร", "group": "หนังสือรับรอง"},
    "CFRD": {"file": "CFRD.docx", "label": "รับรอง รด. ปี 3", "group": "หนังสือรับรอง"},
    "CFRT": {"file": "CFRT.docx", "label": "รับรองผลตรวจเลือก", "group": "หนังสือรับรอง"},
    "NEWRT": {"file": "NEWRT.docx", "label": "ขอใบแทน สด.43", "group": "ขอใบแทน สด.43"},
}
DOC_GROUPS = ["ขอใบแทน สด.8", "หนังสือรับรอง", "ขอใบแทน สด.43"]


def open_path(path: Path):
    try:
        if platform.system() == "Windows":
            os.startfile(str(path))  # noqa
        elif platform.system() == "Darwin":
            subprocess.call(["open", str(path)])
        else:
            subprocess.call(["xdg-open", str(path)])
    except Exception as e:
        messagebox.showwarning("แจ้งเตือน", f"เปิดไฟล์ไม่สำเร็จ: {e}")


class LabelChip(ctk.CTkLabel):
    """ป้ายกำกับสไตล์ 'เซลล์สีเหลือง' ให้ความรู้สึกเหมือนตาราง Excel

    ความกว้างจะขยายอัตโนมัติตามความยาวจริงของข้อความ (วัดจากฟอนต์ที่ใช้จริง)
    เพื่อป้องกันข้อความไทยที่มีสระ/วรรณยุกต์ถูกตัด/ล้นออกนอกกรอบ ("ข้อความตก")
    ค่า width ที่ส่งเข้ามาจะถือเป็นความกว้าง "ขั้นต่ำ" เท่านั้น
    """

    def __init__(self, parent, text, width=110, **kwargs):
        try:
            measured = tkfont.Font(family=UI_FONT_FAMILY, size=16, weight="bold").measure(text)
        except Exception:
            measured = len(text) * 15
        width = max(width, measured + 28)
        super().__init__(
            parent, text=text, width=width, height=34,
            fg_color=COLOR_LABEL_CHIP, text_color=COLOR_LABEL_TEXT,
            corner_radius=0, font=ui_font(16, "bold"),
            **kwargs,
        )


def make_entry(parent, width=140, justify="left", var=None):
    return ctk.CTkEntry(
        parent, width=width, height=34, corner_radius=0, textvariable=var,
        fg_color=COLOR_ENTRY_BG, text_color=COLOR_ENTRY_TEXT,
        border_width=2, border_color=COLOR_ENTRY_BORDER,
        font=ui_font(16),
        justify=justify,
    )


class VarEntry:
    """หน้าตาเหมือน Entry (get/delete/insert) แต่ผูกกับตัวแปรกลาง -> โค้ดเดิมใช้ต่อได้ และทุกหน้าเห็นค่าเดียวกัน"""

    def __init__(self, var):
        self.var = var

    def get(self):
        return self.var.get()

    def delete(self, first=0, last=None):
        self.var.set("")

    def insert(self, index, text):
        self.var.set(str(text) + self.var.get())


def _f(label, key, w, lw=90):
    return ("f", label, key, w, lw)


_PERSON_FULL = ("ข้อมูลบุคคล", [
    [_f("ยศ", "yos", 110, 70), _f("ชื่อ-สกุล", "name", 380)],
    [_f("ปีเกิด (พ.ศ.)", "birth_year", 80, 110), _f("ปีอ้างอิง", "ref_year", 80), ("age",)],
    [_f("สัญชาติ", "nationality", 90), _f("เชื้อชาติ", "ethnicity", 90)],
])
_PERSON_SD8 = ("ข้อมูลบุคคล", [
    [_f("ยศ", "yos", 110, 70), _f("ชื่อ-สกุล", "name", 380)],
    [_f("ปีเกิด (พ.ศ.)", "birth_year", 80, 110)],
])
_PERSON_NAME = ("ข้อมูลบุคคล", [[_f("ชื่อ-สกุล", "name", 460)]])
_PERSON_NAME_YEAR = ("ข้อมูลบุคคล", [
    [_f("ชื่อ-สกุล", "name", 460)],
    [_f("ปีเกิด (พ.ศ.)", "birth_year", 80, 110)],
])
_BIRTHDATE = ("วันเดือนปีเกิด", [
    [_f("เกิดวันที่", "birth_day", 55), ("e", "birth_month", 110), ("e", "birth_yearfull", 80)],
])
_ADDR = ("ที่อยู่ตามภูมิลำเนาทหาร", [
    [_f("บ้านเลขที่", "address_no", 120), _f("หมู่ที่", "moo", 70, 70)],
    [_f("ตำบล", "tambon", 150), _f("อำเภอ", "amphoe", 150)],
    [_f("จังหวัด", "province", 150), _f("เลขบัตร ปชช.", "idcard", 180, 110)],
])
_PARENTS = ("บิดา - มารดา", [
    [_f("ชื่อบิดา", "father_first", 140), _f("สกุลบิดา", "father_last", 140)],
    [_f("ชื่อมารดา", "mother_first", 140), _f("สกุลมารดา", "mother_last", 140)],
])
_MILNO = ("หมายเลขทหาร", [
    [_f("สังกัด,ปี", "tb_prefix", 60, 60), ("e", "tb_year", 80),
     _f("จังหวัด,ลำดับ", "jw_prefix", 60, 70), ("e", "jw_number", 80)],
])
_CFSD = ("รับรองพลทหาร", [
    [_f("สังกัด", "unit", 380)],
    [_f("ผลัดที่", "conscript_batch", 80), _f("ประจำปี", "conscript_year", 100)],
    [_f("ปลดวันที่", "release_day", 55), ("e", "release_month", 110), ("e", "release_year", 80)],
])
_RELEASE = ("วันปลด", [
    [_f("ปลดวันที่", "release_day", 55), ("e", "release_month", 110), ("e", "release_year", 80)],
])
_SD43 = ("ผลตรวจเลือก / สด.43", [
    [_f("วันที่เกณฑ์", "draft_date", 200, 100), _f("บัญชี", "roll_no", 340, 80)],
    [_f("จำนวนนับที่", "count_no", 100, 100), _f("สด.43 เลขที่", "sd43_no", 120, 100)],
    [_f("คนจำพวกที่", "class_type", 80, 100), _f("ส่วนสูง (ซม.)", "height_cm", 80, 100),
     _f("รอบตัว (ซม.)", "chest_cm", 80, 100)],
    [_f("จับสลากได้", "lottery_result", 300, 100), ("hint", '(กรอก "ดำ" หรือ "แดง")')],
])

# ฟอร์มของแต่ละเอกสาร: แสดงเฉพาะช่องที่เอกสารนั้นใช้ (อิงตัวแปรใน templates/*.docx)
DOC_FORMS = {
    "PK14": [_PERSON_FULL, _ADDR, _PARENTS, _MILNO],
    "SD8": [_PERSON_SD8, _ADDR, _PARENTS, _MILNO],
    "CFSD": [_PERSON_NAME, _BIRTHDATE, _ADDR, _PARENTS, _CFSD],
    "CFRD": [_PERSON_NAME, _BIRTHDATE, _ADDR, _PARENTS, _RELEASE],
    "CFRT": [_PERSON_NAME, _BIRTHDATE, _ADDR, _PARENTS, _SD43],
    "NEWRT": [_PERSON_NAME_YEAR, _ADDR, _PARENTS, _SD43],
}


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        global UI_FONT_FAMILY
        UI_FONT_FAMILY = _detect_font(self)

        self.title(f"ระบบขอ สด.8 ใหม่ / ป.ค.14 / หนังสือรับรอง  [Beta {APP_VER}]")
        self.geometry("1420x860")
        self.minsize(900, 560)
        self.configure(fg_color=COLOR_BG)
        try:
            self.iconbitmap(str(BASE_DIR / "favicon.ico"))
        except Exception:
            pass

        self.status_var = ctk.StringVar(value="พร้อมใช้งาน")
        # ตัวแปรกลางของทุกช่อง (ผูกกับทุกหน้าฟอร์ม) + ตัวห่อที่หน้าตาเหมือน Entry ให้โค้ดเดิมใช้ต่อได้
        self.vars = {k: ctk.StringVar() for k in ALL_FIELD_KEYS}
        self.entries = {k: VarEntry(v) for k, v in self.vars.items()}
        self.age_labels = []
        self.vars["nationality"].set("ไทย")
        self.vars["ethnicity"].set("ไทย")
        self.vars["ref_year"].set(str(datetime.now().year + 543))
        self.vars["tb_prefix"].set("ท.บ.")
        self.vars["jw_prefix"].set("ส.น.")
        for _k in ("birth_year", "ref_year"):
            self.vars[_k].trace_add("write", lambda *a: self._update_age())

        self._build_header()
        self._build_workspace()
        self._build_status_bar()

        self._load_autosave()
        self._schedule_autosave()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # ตั้งค่าหน่วยงาน: ถ้ายังไม่เคยตั้งค่า (หรือข้อมูลจำเป็นไม่ครบ) เปิดหน้าตั้งค่าอัตโนมัติ
        self.org = org_settings.load()
        if not org_settings.is_configured():
            self.after(500, lambda: self._open_settings(first_run=True))
        self.after(900, self._check_templates)
        self._enable_thai_shortcuts()
        self.after(3000, lambda: self._check_update(silent=True))

    # ------------------------------------------------------------
    # UI CONSTRUCTION
    # ------------------------------------------------------------
    def _build_header(self):
        header = ctk.CTkFrame(
            self, fg_color=COLOR_SURFACE, corner_radius=0,
            border_width=2, border_color=COLOR_BORDER,
        )
        header.pack(fill="x", padx=20, pady=(18, 10))

        logo_label = None
        if LOGO_PATH.exists():
            try:
                logo_img = Image.open(LOGO_PATH).convert("RGBA")
                if LOGO_PATH.name == "sd8_logo.png":  # pixel-art: ขยายแบบ nearest กันภาพเบลอ
                    logo_img = logo_img.resize((184, 184), Image.NEAREST)
                self.logo_image = ctk.CTkImage(
                    light_image=logo_img, dark_image=logo_img, size=(92, 92)
                )
                logo_label = ctk.CTkLabel(header, text="", image=self.logo_image)
                logo_label.pack(side="left", padx=(18, 14), pady=12)
            except Exception:
                logo_label = None

        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.pack(side="left", fill="y", pady=16)
        ctk.CTkLabel(
            title_box, text="  AMGF1  |  DOCUMENT CONTROL จัดทำโดย ซานต้าคลอสไอลูกไม่รักดี",
            font=ui_font(13, "bold"), text_color=COLOR_GOLD, anchor="w",
        ).pack(anchor="w")
        ctk.CTkLabel(
            title_box, text="  ระบบจัดทำเอกสาร สด.8 / ป.ค.14 / หนังสือรับรอง",
            font=ui_font(25, "bold"), text_color=COLOR_TEXT_LIGHT, anchor="w",
        ).pack(anchor="w", pady=(2, 0))
        ctk.CTkLabel(
            title_box, text="  เลือกเอกสารทางซ้าย กรอกฟอร์ม (หรือดึงจากเว็บ) แล้วกดสร้างด้านล่าง",
            font=ui_font(13), text_color=COLOR_TEXT_MUTED, anchor="w",
        ).pack(anchor="w", pady=(2, 0))

        right = ctk.CTkFrame(header, fg_color="transparent")
        right.pack(side="right", padx=16, pady=16)
        ctk.CTkButton(
            right, text="ล้างข้อมูล", width=120, height=38,
            fg_color="#4A1020", hover_color="#6A1830", text_color="#FF4D6D",
            corner_radius=0, font=ui_font(14, "bold"),
            command=self._clear_all,
        ).pack()
        ctk.CTkButton(
            right, text="ตั้งค่าหน่วยงาน", width=120, height=38,
            fg_color=COLOR_BTN_SECONDARY, hover_color=COLOR_BTN_TERTIARY, text_color=COLOR_GOLD,
            corner_radius=0, font=ui_font(14, "bold"),
            command=self._open_settings,
        ).pack(pady=(8, 0))
        ctk.CTkButton(
            right, text="ตรวจอัปเดต", width=120, height=38,
            fg_color=COLOR_BTN_SECONDARY, hover_color=COLOR_BTN_TERTIARY, text_color=COLOR_GOLD,
            corner_radius=0, font=ui_font(14, "bold"),
            command=lambda: self._check_update(silent=False),
        ).pack(pady=(8, 0))

    # ------------------------------------------------------------
    # โครงหน้าจอ: ซ้าย = เลือกเอกสาร / ขวา = ฟอร์มของเอกสารนั้น
    # ทุกช่องใช้ตัวแปรกลางชุดเดียว (self.vars) -> ค่าเหมือนกันใช้ร่วมทุกเอกสาร แก้ที่ไหนเปลี่ยนทุกที่
    # ------------------------------------------------------------
    def _build_workspace(self):
        self.workspace = ctk.CTkFrame(self, fg_color="transparent")
        self.workspace.pack(fill="both", expand=True, padx=20, pady=(0, 14))
        self.workspace.grid_columnconfigure(0, weight=0)
        self.workspace.grid_columnconfigure(1, weight=1)
        self.workspace.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_form_area()
        self._select_doc("PK14")

    def _row(self, parent, r, pady=6):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.grid(row=r, column=0, sticky="ew", pady=pady, padx=4)
        return f

    def _section_title(self, parent, r, text):
        lbl = ctk.CTkLabel(
            parent, text=text, font=ui_font(15, "bold"),
            text_color=COLOR_GOLD, anchor="w",
        )
        lbl.grid(row=r, column=0, sticky="w", padx=4, pady=(14, 2))

    # ---------- ซ้าย: เลือกเอกสาร ----------
    def _build_sidebar(self):
        panel = ctk.CTkFrame(
            self.workspace, width=250, fg_color=COLOR_SURFACE,
            corner_radius=0, border_width=2, border_color=COLOR_BORDER,
        )
        panel.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        panel.grid_propagate(False)
        panel.pack_propagate(False)

        ctk.CTkLabel(panel, text="DOCUMENTS", font=pixel_font(13), text_color=COLOR_GOLD,
                     anchor="w").pack(anchor="w", padx=16, pady=(16, 0))
        ctk.CTkLabel(panel, text="เลือกเอกสารที่ต้องการกรอก", font=ui_font(13),
                     text_color=COLOR_TEXT_MUTED, anchor="w").pack(anchor="w", padx=16, pady=(2, 10))

        self.web_btn = ctk.CTkButton(
            panel, text="ดึงข้อมูลจากเว็บ TDC", height=44,
            fg_color=COLOR_BTN_SECONDARY, hover_color=COLOR_BTN_TERTIARY,
            text_color=COLOR_GOLD, corner_radius=0, font=ui_font(14, "bold"),
            command=self._pull_from_web,
        )
        self.web_btn.pack(fill="x", padx=14, pady=(0, 8))
        ctk.CTkLabel(panel, text="ดึงครั้งเดียว ข้อมูลเติมให้ทุกเอกสาร", font=ui_font(12),
                     text_color=COLOR_TEXT_MUTED, anchor="w").pack(anchor="w", padx=16, pady=(0, 6))
        ctk.CTkFrame(panel, height=1, fg_color=COLOR_BORDER).pack(fill="x", padx=14, pady=(2, 10))

        self.doc_buttons = {}
        for group in DOC_GROUPS:
            title = group
            ctk.CTkLabel(panel, text=title, font=ui_font(13, "bold"),
                         text_color=COLOR_TEXT_MUTED, anchor="w").pack(anchor="w", padx=16, pady=(4, 4))
            for key, info in DOCS.items():
                if info["group"] != group:
                    continue
                btn = ctk.CTkButton(
                    panel, text=info["label"], height=46, corner_radius=0, anchor="w",
                    fg_color=COLOR_BTN_TERTIARY, hover_color="#2E2E66",
                    text_color=COLOR_TEXT_LIGHT, font=ui_font(15, "bold"),
                    command=lambda k=key: self._select_doc(k),
                )
                btn.pack(fill="x", padx=14, pady=3)
                self.doc_buttons[key] = btn

    # ---------- ขวา: ฟอร์ม + ปุ่มสร้าง ----------
    def _build_form_area(self):
        area = ctk.CTkFrame(self.workspace, fg_color=COLOR_CARD, corner_radius=0,
                            border_width=2, border_color=COLOR_BORDER)
        area.grid(row=0, column=1, sticky="nsew")

        head = ctk.CTkFrame(area, fg_color=COLOR_PANEL_HEADER, corner_radius=0)
        head.pack(fill="x")
        self.form_title = ctk.CTkLabel(head, text="", font=ui_font(18, "bold"),
                                       text_color=COLOR_GOLD, anchor="w")
        self.form_title.pack(side="left", padx=16, pady=10)
        ctk.CTkLabel(head, text="ค่าที่ซ้ำกันใช้ร่วมกันทุกเอกสาร แก้ที่ไหนเปลี่ยนทุกที่",
                     font=ui_font(12), text_color=COLOR_TEXT_MUTED).pack(side="right", padx=16)

        bar = ctk.CTkFrame(area, fg_color=COLOR_SURFACE, corner_radius=0)
        bar.pack(side="bottom", fill="x")
        self.make_btn = ctk.CTkButton(
            bar, text="สร้างเอกสารนี้", height=46, width=200, corner_radius=0,
            fg_color=COLOR_BTN_PRIMARY, hover_color=COLOR_GOLD_HOVER, text_color="#0B0B1A",
            font=ui_font(15, "bold"), command=lambda: self._generate(self.current_doc),
        )
        self.make_btn.pack(side="left", padx=(14, 8), pady=10)
        ctk.CTkButton(
            bar, text="สร้างทั้งหมด (6 ไฟล์)", height=46, corner_radius=0,
            fg_color=COLOR_BTN_TERTIARY, hover_color="#2E2E66", text_color=COLOR_TEXT_LIGHT,
            font=ui_font(14, "bold"), command=self._generate_all,
        ).pack(side="left", padx=8, pady=10)
        ctk.CTkButton(
            bar, text="เปิดโฟลเดอร์ Export", height=46, corner_radius=0,
            fg_color=COLOR_BTN_SECONDARY, hover_color=COLOR_BTN_TERTIARY, text_color=COLOR_SILVER,
            font=ui_font(14, "bold"), command=lambda: open_path(OUTPUT_DIR),
        ).pack(side="right", padx=14, pady=10)

        self.form_host = ctk.CTkFrame(area, fg_color="transparent", corner_radius=0)
        self.form_host.pack(fill="both", expand=True, padx=8, pady=8)
        self.pages = {}
        self.current_doc = None

    def _select_doc(self, key):
        if self.current_doc and self.current_doc in self.pages:
            self.pages[self.current_doc].pack_forget()
        if key not in self.pages:
            self.pages[key] = self._build_page(key)
        self.pages[key].pack(fill="both", expand=True)
        self.current_doc = key
        for k, b in self.doc_buttons.items():
            on = k == key
            b.configure(fg_color=COLOR_GOLD if on else COLOR_BTN_TERTIARY,
                        text_color="#0B0B1A" if on else COLOR_TEXT_LIGHT,
                        hover_color=COLOR_GOLD_HOVER if on else "#2E2E66")
        self.form_title.configure(text=f"ฟอร์ม: {DOCS[key]['label']}")
        self.make_btn.configure(text=f"สร้าง {DOCS[key]['label']}")
        self._update_age()

    def _build_page(self, doc_key):
        page = ctk.CTkScrollableFrame(
            self.form_host, fg_color="transparent",
            scrollbar_button_color=COLOR_BTN_TERTIARY,
            scrollbar_button_hover_color=COLOR_GOLD,
        )
        page.grid_columnconfigure(0, weight=1)
        r = 0
        for title, rows in DOC_FORMS[doc_key]:
            self._section_title(page, r, title); r += 1
            for spec_row in rows:
                row = self._row(page, r); r += 1
                for i, spec in enumerate(spec_row):
                    last = i == len(spec_row) - 1
                    pad = (0, 0) if last else (0, 10)
                    if spec[0] == "f":
                        _, label, key, w, lw = spec
                        LabelChip(row, label, width=lw).pack(side="left", padx=(0, 4))
                        self._entry(row, key, w).pack(side="left", padx=pad)
                    elif spec[0] == "e":
                        _, key, w = spec
                        self._entry(row, key, w).pack(side="left", padx=(0, 6) if not last else (0, 0))
                    elif spec[0] == "age":
                        LabelChip(row, "อายุ", width=60).pack(side="left", padx=(0, 4))
                        lbl = ctk.CTkLabel(
                            row, text="-", width=60, height=34, fg_color=COLOR_LABEL_CHIP,
                            text_color=COLOR_LABEL_TEXT, corner_radius=0, font=ui_font(16, "bold"),
                        )
                        lbl.pack(side="left")
                        self.age_labels.append(lbl)
                    elif spec[0] == "hint":
                        ctk.CTkLabel(row, text=spec[1], font=ui_font(13),
                                     text_color=COLOR_TEXT_MUTED).pack(side="left", padx=(10, 0))
        return page

    def _entry(self, parent, key, width):
        return make_entry(parent, width=width, var=self.vars[key])

    def _build_status_bar(self):
        bar = ctk.CTkFrame(self, fg_color=COLOR_PANEL_HEADER, height=34, corner_radius=0)
        bar.pack(fill="x", side="bottom")
        ctk.CTkLabel(
            bar, textvariable=self.status_var, text_color=COLOR_TEXT_MUTED,
            font=ui_font(12), anchor="w",
        ).pack(side="left", padx=16, pady=4)

    # ------------------------------------------------------------
    # DATA HANDLING
    # ------------------------------------------------------------
    def _collect_data(self) -> dict:
        return {k: v.get().strip() for k, v in self.entries.items()}

    def _update_age(self):
        data = self._collect_data()
        from utils import compute_age
        age = compute_age(data.get("birth_year", ""), data.get("ref_year", ""))
        for lbl in getattr(self, "age_labels", []):
            lbl.configure(text=age if age else "-")

    def _load_autosave(self):
        data = storage.load_autosave()
        for k, entry in self.entries.items():
            if k in data and data[k]:
                entry.delete(0, "end")
                entry.insert(0, data[k])
        self._update_age()

    def _schedule_autosave(self):
        storage.autosave(self._collect_data())
        self.after(4000, self._schedule_autosave)

    def _clear_all(self):
        if not messagebox.askyesno("ยืนยัน", "ต้องการล้างข้อมูลทั้งหมดในฟอร์มหรือไม่?"):
            return
        for k, entry in self.entries.items():
            entry.delete(0, "end")
        self.entries["nationality"].insert(0, "ไทย")
        self.entries["ethnicity"].insert(0, "ไทย")
        self.entries["ref_year"].insert(0, str(datetime.now().year + 543))
        self.entries["tb_prefix"].insert(0, "ทบ.")
        self.entries["jw_prefix"].insert(0, "ส.น.")
        self._update_age()
        storage.autosave(self._collect_data())
        self.status_var.set("ล้างข้อมูลแล้ว")

    # ------------------------------------------------------------
    # ดึงข้อมูลจากเว็บ TDC Recruit (หน้า ตรวจสอบข้อมูลบุคคล)
    # ------------------------------------------------------------
    def _pull_from_web(self):
        try:
            import web_import  # noqa: F401  (ตรวจว่า playwright พร้อมใช้)
        except Exception as e:
            messagebox.showerror("ใช้ฟีเจอร์นี้ไม่ได้", f"ขาดไลบรารี playwright หรือโหลดไม่สำเร็จ\n\n{e}\n\n"
                                 "ติดตั้ง: pip install playwright  แล้ว  playwright install chromium")
            return

        # ถามเลขบัตรก่อนเสมอ แล้วค่อยเปิดเบราว์เซอร์/เข้าสู่ระบบ/ดึงข้อมูล
        IdPromptDialog(
            self, self._theme(), ui_font, self.entries["idcard"].get(),
            on_ok=self._start_pull,
            on_cancel=lambda: self.status_var.set("ยกเลิกการดึงข้อมูล"),
        )

    def _theme(self):
        return {
            "bg": COLOR_BG, "surface": COLOR_SURFACE, "text": COLOR_TEXT_LIGHT, "muted": COLOR_TEXT_MUTED,
            "accent": COLOR_GOLD, "accent_hover": COLOR_GOLD_HOVER, "btn2": COLOR_BTN_SECONDARY,
            "btn3": COLOR_BTN_TERTIARY, "entry_bg": COLOR_ENTRY_BG, "entry_border": COLOR_ENTRY_BORDER,
            "entry_text": COLOR_ENTRY_TEXT,
        }

    def _start_pull(self, cid):
        from web_import import WebSession
        self._pending_cid = cid
        self.vars["idcard"].set(cid)
        sess = getattr(self, "_web", None)
        if sess is not None and sess.is_open and sess.ready.is_set() and not sess.error:
            self._ask_and_fetch()
            return

        self._ensure_linkage()
        # เปิดเบราว์เซอร์ (ครั้งแรกอาจต้องโหลด Chromium 1-3 นาที) แล้วรอจนพร้อมโดยไม่ค้างหน้าจอ
        self._web = WebSession(log=lambda m: self.after(0, lambda: self.status_var.set(m)),
                              unit_hint=org_settings.load().get("amphoe", ""))
        self._web.open()
        self.web_btn.configure(state="disabled")
        self.status_var.set("กำลังเตรียมเบราว์เซอร์ (ครั้งแรกอาจนาน 1-3 นาที ต้องต่อเน็ต) กรุณารอ...")
        self._wait_web_ready()

    def _ensure_linkage(self):
        """เช็กว่า Linkage2Client เปิดอยู่ไหม ถ้าไม่เปิดก็เปิดให้ (ไม่ขวางการดึงข้อมูลถ้าทำไม่ได้)"""
        st, msg = linkage_check.ensure_running()
        if st == "notfound":
            from tkinter import filedialog
            if messagebox.askyesno("Linkage2Client", msg + "\n\nเลือกไฟล์ .exe ของโปรแกรมเองตอนนี้ (จำไว้ใช้ครั้งต่อไป)?"):
                p = filedialog.askopenfilename(title="เลือก Linkage2Client .exe", filetypes=[("โปรแกรม", "*.exe")])
                if p:
                    linkage_check.save_path(p)
                    st, msg = linkage_check.ensure_running()
        self.status_var.set(f"Linkage2Client: {msg}")

    def _wait_web_ready(self):
        s = self._web
        if not s.ready.is_set():
            self.after(500, self._wait_web_ready)
            return
        self.web_btn.configure(state="normal")
        if s.error:
            self.status_var.set("เปิดเบราว์เซอร์ไม่สำเร็จ")
            messagebox.showerror("เปิดเบราว์เซอร์ไม่สำเร็จ", s.error)
            return
        self.status_var.set("กำลังเข้าสู่ระบบ Recruit ด้วยโทเค็น Linkage2Client... (ถ้าต้องสแกน ThaiD ให้สแกนที่หน้าที่เปิด)")
        self._wait_login()

    def _wait_login(self):
        s = self._web
        if not s.login_done.is_set():
            self.after(500, self._wait_login)
            return
        if not s.login_ok:
            messagebox.showinfo(
                "Login เข้า TDC Recruit",
                "เข้าสู่ระบบอัตโนมัติไม่สำเร็จ\n"
                "ในหน้าต่างเบราว์เซอร์ที่เปิดขึ้น ให้ Login เองให้เสร็จ\n"
                "แล้วกด OK ที่นี่ โปรแกรมจะค้นหาและดึงข้อมูลให้ทันที",
            )
        self._ask_and_fetch()

    def _ask_and_fetch(self):
        cid = getattr(self, "_pending_cid", "") or "".join(ch for ch in self.entries["idcard"].get() if ch.isdigit())
        self.web_btn.configure(state="disabled")
        self.status_var.set(f"กำลังค้นหาและดึงข้อมูล {cid} จากเว็บ...")

        def work():
            try:
                res = self._web.fetch(cid)
                self.after(0, lambda: self._apply_web_result(*res))
            except Exception as e:
                err = str(e) or repr(e)  # e ถูกลบหลังออกจาก except -> เก็บข้อความไว้ก่อน
                self.after(0, lambda: self._web_failed(err))

        threading.Thread(target=work, daemon=True).start()

    def _web_failed(self, msg):
        self.web_btn.configure(state="normal")
        self.status_var.set("ดึงข้อมูลไม่สำเร็จ")
        messagebox.showerror("ดึงข้อมูลจากเว็บไม่สำเร็จ", msg)

    def _apply_web_result(self, values, notes):
        self.web_btn.configure(state="normal")
        # ช่องที่มีค่าอยู่แล้วและจะถูกเปลี่ยน -> ถามก่อนเขียนทับ
        changed = [k for k, v in values.items()
                   if k in self.entries and self.entries[k].get().strip() not in ("", v)]
        overwrite = True
        if changed:
            overwrite = messagebox.askyesno(
                "เขียนทับข้อมูลเดิม?",
                f"มี {len(changed)} ช่องที่มีค่าอยู่แล้วและต่างจากเว็บ\n"
                "กด 'ใช่' = เขียนทับด้วยค่าจากเว็บ\nกด 'ไม่' = เติมเฉพาะช่องที่ว่าง",
            )
        filled = 0
        for k, v in values.items():
            e = self.entries.get(k)
            if e is None:
                continue
            if e.get().strip() and not overwrite:
                continue
            e.delete(0, "end")
            e.insert(0, v)
            filled += 1
        self._update_age()
        storage.autosave(self._collect_data())
        msg = f"ดึงจากเว็บแล้ว {filled} ช่อง"
        if notes:
            msg += " | " + "; ".join(notes)
        self.status_var.set(msg)

    # ------------------------------------------------------------
    # คีย์ลัดตอนใช้แป้นพิมพ์ภาษาไทย: Ctrl+แ = คัดลอก, Ctrl+อ = วาง, Ctrl+ฉ = ตัด, Ctrl+ฟ = เลือกทั้งหมด
    # (Tk ผูกคีย์ลัดกับตัวอักษรอังกฤษ พอสลับเป็นแป้นไทยจึงไม่ทำงาน -> ใช้รหัสปุ่มจริงของ Windows แทน)
    # ------------------------------------------------------------
    def _enable_thai_shortcuts(self):
        if platform.system() == "Windows":
            self.bind_all("<Control-Key>", self._on_ctrl_key, add="+")

    @staticmethod
    def _on_ctrl_key(event):
        if str(event.keysym).lower() in ("c", "v", "x", "a"):
            return None  # แป้นอังกฤษ: Tk จัดการเองอยู่แล้ว
        w = event.widget
        virtual = {67: "<<Copy>>", 86: "<<Paste>>", 88: "<<Cut>>"}.get(event.keycode)
        try:
            if virtual:
                w.event_generate(virtual)
                return "break"
            if event.keycode == 65:  # A
                if hasattr(w, "select_range"):
                    w.select_range(0, "end")
                    w.icursor("end")
                elif hasattr(w, "tag_add"):
                    w.tag_add("sel", "1.0", "end-1c")
                return "break"
        except Exception:
            pass
        return None

    # ------------------------------------------------------------
    # อัปเดตจาก GitHub (แจ้งเท่านั้น ต้องกดยืนยันเอง)
    # ------------------------------------------------------------
    def _check_update(self, silent=True, url=None):
        def work():
            try:
                info = update_check.check(APP_VER, url or update_check.DEFAULT_URL)
                self.after(0, lambda: self._update_result(info, silent))
            except Exception as e:
                err = str(e) or repr(e)
                if not silent:
                    self.after(0, lambda: self._update_failed(err))
        threading.Thread(target=work, daemon=True).start()

    def _update_failed(self, err):
        import tkinter.simpledialog as sd
        new = sd.askstring("ตรวจอัปเดตไม่สำเร็จ",
                           f"{err}\n\nวางลิงก์ update.json ใหม่ (ถ้ามี) แล้วกด OK:", parent=self)
        if new and new.strip():
            self._check_update(silent=False, url=new.strip())

    def _update_result(self, info, silent):
        if not info:
            if not silent:
                messagebox.showinfo("อัปเดต", f"เป็นรุ่นล่าสุดแล้ว ({APP_VER})")
            return
        if not messagebox.askyesno(
                "มีรุ่นใหม่",
                f"รุ่นปัจจุบัน {APP_VER} -> รุ่นใหม่ {info['version']}\n"
                f"ไฟล์: {info.get('file','SD8System_Setup.exe')} ({int(info.get('size',0))//1024} KB)\n\n"
                "ดาวน์โหลดและติดตั้งตอนนี้? (ข้อมูล/ตั้งค่าเดิมไม่ถูกลบ)"):
            return
        self.status_var.set("กำลังดาวน์โหลดอัปเดต...")

        def work():
            try:
                path = update_check.download(
                    info, lambda g, t: self.after(0, lambda: self.status_var.set(f"ดาวน์โหลดอัปเดต {g*100//t}%")))
                self.after(0, lambda: self._run_installer(path))
            except Exception as e:
                err = str(e) or repr(e)
                self.after(0, lambda: (self.status_var.set("อัปเดตไม่สำเร็จ"),
                                       messagebox.showerror("อัปเดตไม่สำเร็จ", err)))
        threading.Thread(target=work, daemon=True).start()

    def _run_installer(self, path):
        if not messagebox.askyesno("พร้อมติดตั้ง", "ตรวจ SHA-256 ผ่านแล้ว\nโปรแกรมจะปิดตัวเองและเปิดตัวติดตั้ง ดำเนินการต่อ?"):
            return
        try:
            os.startfile(path)  # Windows
        except Exception as e:
            messagebox.showerror("เปิดตัวติดตั้งไม่ได้", f"{e}\nไฟล์อยู่ที่: {path}")
            return
        self.destroy()

    def _check_templates(self):
        """เตือนถ้าเทมเพลตเป็นรุ่นเก่า (ไม่มีช่องวันที่อัตโนมัติ/ช่องหน่วยงาน) ซึ่งทำให้ระบบตั้งค่าไม่มีผล"""
        import zipfile
        old = []
        for key, info in DOCS.items():
            p = TEMPLATE_DIR / info["file"]
            try:
                with zipfile.ZipFile(p) as z:
                    x = z.read("word/document.xml").decode("utf-8", "ignore")
                if "เดือนเต็ม" not in x and "เดือนย่อ" not in x:
                    old.append(info["file"])
            except Exception:
                old.append(info["file"] + " (เปิดไม่ได้/ไม่มีไฟล์)")
        if old:
            messagebox.showwarning(
                "เทมเพลตเป็นรุ่นเก่า",
                "ไฟล์เทมเพลตต่อไปนี้ยังไม่ใช่รุ่นที่รองรับวันที่อัตโนมัติและการตั้งค่าหน่วยงาน:\n- "
                + "\n- ".join(old)
                + f"\n\nโปรแกรมอ่านเทมเพลตจากโฟลเดอร์:\n{TEMPLATE_DIR}\n\n"
                "ให้นำไฟล์ในโฟลเดอร์ templates จากชุดล่าสุดไปวางทับที่นี่",
            )

    def _leftover_placeholders(self, path):
        import re
        import zipfile
        try:
            with zipfile.ZipFile(path) as z:
                x = z.read("word/document.xml").decode("utf-8", "ignore")
        except Exception:
            return []
        return sorted(set(re.findall(r"&lt;&lt;([^&]+)&gt;&gt;", x)))

    # ------------------------------------------------------------
    # ตั้งค่าหน่วยงาน
    # ------------------------------------------------------------
    def _open_settings(self, first_run=False):
        if getattr(self, "_settings_win", None) is not None:
            try:
                if self._settings_win.winfo_exists():
                    self._settings_win.focus_force()
                    return
            except Exception:
                pass
        theme = self._theme()
        self._settings_win = SettingsDialog(
            self, theme, ui_font, self.org, self._on_settings_saved,
            first_run=first_run,
            on_cancel=lambda: self.status_var.set("ยังไม่ได้ตั้งค่าหน่วยงาน — กดปุ่ม 'ตั้งค่าหน่วยงาน' ก่อนสร้างเอกสาร"),
        )

    def _on_settings_saved(self, values, path):
        self.org = org_settings.load()
        # ตัวย่อจังหวัดใน "เลขเครื่องหมาย (จว.)" ใช้ค่าเดียวกัน ถ้ายังเป็นค่าเริ่มต้น/ว่าง
        jw = self.entries.get("jw_prefix")
        if jw is not None and jw.get().strip() in ("", "ส.น."):
            jw.delete(0, "end")
            jw.insert(0, self.org.get("prov_abbr", "") or "ส.น.")
        self.status_var.set(f"บันทึกการตั้งค่าหน่วยงานแล้ว: {path}")

    # ------------------------------------------------------------
    # DOCUMENT GENERATION
    # ------------------------------------------------------------
    def _generate(self, doc_key: str, quiet: bool = False):
        info = DOCS[doc_key]
        template_path = TEMPLATE_DIR / info["file"]

        if not template_path.exists():
            messagebox.showerror(
                "ไม่พบไฟล์เทมเพลต",
                f"ไม่พบไฟล์ {info['file']} ในโฟลเดอร์:\n{TEMPLATE_DIR}\n\n"
                f"กรุณานำไฟล์ .docx ต้นแบบไปวางในโฟลเดอร์ templates ก่อนใช้งาน",
            )
            return None

        if not org_settings.is_configured():
            messagebox.showinfo("ยังไม่ได้ตั้งค่าหน่วยงาน",
                                "กรุณาตั้งค่าหน่วยงาน/ผู้ลงนามก่อนสร้างเอกสาร")
            self._open_settings(first_run=True)
            return None

        data = self._collect_data()
        self._update_age()
        data["ref_year"] = self.entries["ref_year"].get().strip()
        mapping = build_mapping(data)
        mapping.update(org_settings.org_mapping(self.org))

        name_for_file = data.get("name", "").strip() or "ไม่ระบุชื่อ"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = "".join(c for c in name_for_file if c not in '\\/:*?"<>|')
        output_path = OUTPUT_DIR / f"{doc_key}_{safe_name}_{timestamp}.docx"

        try:
            fill_docx(str(template_path), str(output_path), mapping)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถสร้างเอกสารได้:\n{e}")
            return None

        left = self._leftover_placeholders(output_path)
        if left and not quiet:
            messagebox.showwarning(
                "มีช่องที่ไม่ได้เติม",
                "ในเอกสารยังมีช่องที่ระบบไม่รู้จัก (อาจใช้โปรแกรมกับเทมเพลตคนละรุ่น):\n- "
                + "\n- ".join(left[:12]),
            )

        storage.autosave(data)
        try:
            storage.log_to_database(info["label"], data)
        except Exception as e:
            print("บันทึกฐานข้อมูลล้มเหลว:", e)

        self.status_var.set(f"สร้างไฟล์ {output_path.name} เรียบร้อยแล้ว")

        if quiet:
            return output_path
        if messagebox.askyesno("สำเร็จ", f"สร้างไฟล์ {info['label']} เรียบร้อยแล้ว\nต้องการเปิดไฟล์เลยหรือไม่?"):
            open_path(output_path)
        return output_path

    def _generate_all(self):
        if not org_settings.is_configured():
            messagebox.showinfo("ยังไม่ได้ตั้งค่าหน่วยงาน", "กรุณาตั้งค่าหน่วยงาน/ผู้ลงนามก่อนสร้างเอกสาร")
            self._open_settings(first_run=True)
            return
        done, failed = [], []
        for key in DOCS:
            p = self._generate(key, quiet=True)
            (done if p else failed).append(DOCS[key]["label"])
        if done and messagebox.askyesno(
                "สร้างทั้งหมดเสร็จ",
                f"สร้างแล้ว {len(done)} ไฟล์: " + ", ".join(done)
                + (f"\nไม่สำเร็จ: {', '.join(failed)}" if failed else "")
                + "\n\nเปิดโฟลเดอร์ Export เลยหรือไม่?"):
            open_path(OUTPUT_DIR)

    def _on_close(self):
        storage.autosave(self._collect_data())
        self.destroy()


def main():
    try:
        app = App()
        app.mainloop()
    except Exception:
        log_path = _write_crash_log("RUNTIME ERROR")
        try:
            messagebox.showerror(
                "เกิดข้อผิดพลาด - โปรแกรมต้องปิดตัว",
                "เกิดข้อผิดพลาดที่ทำให้โปรแกรมทำงานต่อไม่ได้:\n\n"
                + traceback.format_exc(limit=3)
                + (f"\n\nรายละเอียดทั้งหมดถูกบันทึกไว้ที่:\n{log_path}" if log_path else ""),
            )
        except Exception:
            pass
        raise


if __name__ == "__main__":
    main()
