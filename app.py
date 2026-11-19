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
    from settings_dialog import SettingsDialog
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

APP_VER = "0.9.0"  # Beta — ต้องตรงกับ installer.iss (AppVersion) และ update.json
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
    "PK14": {"file": "PK14.docx", "label": "ป.ค.14", "group": "หลัก"},
    "SD8": {"file": "SD8.docx", "label": "ปะหน้า (สด.8)", "group": "หลัก"},
    "CFSD": {"file": "CFSD.docx", "label": "รับรองพลทหาร", "group": "หลัก"},
    "CFRD": {"file": "CFRD.docx", "label": "รับรอง รด. ปี 3", "group": "หลัก"},
    "CFRT": {"file": "CFRT.docx", "label": "รับรองผลตรวจเลือก", "group": "สด.43"},
    "NEWRT": {"file": "NEWRT.docx", "label": "ขอใบแทน สด.43", "group": "สด.43"},
}


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


def make_entry(parent, width=140, justify="left"):
    return ctk.CTkEntry(
        parent, width=width, height=34, corner_radius=0,
        fg_color=COLOR_ENTRY_BG, text_color=COLOR_ENTRY_TEXT,
        border_width=2, border_color=COLOR_ENTRY_BORDER,
        font=ui_font(16),
        justify=justify,
    )


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

        self.entries = {}
        self.status_var = ctk.StringVar(value="พร้อมใช้งาน")

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
            title_box, text="  กรอกข้อมูลครั้งเดียว แล้วเลือก Export เอกสารที่ต้องการจากแถบด้านขวา",
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

    def _build_workspace(self):
        self.workspace = ctk.CTkFrame(self, fg_color="transparent")
        self.workspace.pack(fill="both", expand=True, padx=20, pady=(0, 14))
        self.workspace.grid_columnconfigure(0, weight=1)
        self.workspace.grid_columnconfigure(1, weight=0)
        self.workspace.grid_rowconfigure(0, weight=1)

        self._build_tabs()
        self._build_button_bar()

    def _build_tabs(self):
        self.tabview = ctk.CTkTabview(
            self.workspace, fg_color=COLOR_CARD, border_width=2, border_color=COLOR_BORDER,
            segmented_button_fg_color=COLOR_PANEL_HEADER,
            segmented_button_selected_color=COLOR_TAB_SELECTED,
            segmented_button_unselected_color=COLOR_BTN_SECONDARY,
            segmented_button_selected_hover_color=COLOR_TAB_HOVER,
            text_color=COLOR_TEXT_LIGHT, corner_radius=0,
        )
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)

        tab_personal = self.tabview.add("ข้อมูลส่วนตัว")
        tab_conscript = self.tabview.add("ตรวจเลือกทหาร / รด.")
        tab_sd43 = self.tabview.add("ใบแทน สด.43")

        self._build_personal_tab(tab_personal)
        self._build_conscript_tab(tab_conscript)
        self._build_sd43_tab(tab_sd43)

        try:
            self.tabview._segmented_button.configure(font=ui_font(16, "bold"))
        except Exception:
            pass

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

    # ---------------- TAB 1: ข้อมูลส่วนตัว ----------------
    def _build_personal_tab(self, tab):
        # ห่อเนื้อหาด้วย ScrollableFrame เสมอ - ป้องกันฟิลด์ท้ายๆ (เช่น ทบ./ส.น.)
        # ถูกตัดหายเวลาหน้าจอ/หน้าต่างเตี้ยกว่าความสูงของเนื้อหาทั้งหมด
        parent = ctk.CTkScrollableFrame(
            tab, fg_color="transparent",
            scrollbar_button_color=COLOR_BTN_TERTIARY,
            scrollbar_button_hover_color=COLOR_GOLD,
        )
        parent.pack(fill="both", expand=True)
        parent.grid_columnconfigure(0, weight=1)
        r = 0

        row = self._row(parent, r); r += 1
        LabelChip(row, "ยศ", width=70).pack(side="left", padx=(0, 4))
        self.entries["yos"] = make_entry(row, width=110)
        self.entries["yos"].pack(side="left", padx=(0, 10))
        LabelChip(row, "ชื่อ-สกุล", width=90).pack(side="left", padx=(0, 4))
        self.entries["name"] = make_entry(row, width=380)
        self.entries["name"].pack(side="left")

        row = self._row(parent, r); r += 1
        LabelChip(row, "ปีเกิด (พ.ศ.)", width=110).pack(side="left", padx=(0, 4))
        self.entries["birth_year"] = make_entry(row, width=80)
        self.entries["birth_year"].pack(side="left", padx=(0, 10))
        LabelChip(row, "ปีอ้างอิง", width=90).pack(side="left", padx=(0, 4))
        self.entries["ref_year"] = make_entry(row, width=80)
        self.entries["ref_year"].pack(side="left", padx=(0, 10))
        self.entries["ref_year"].insert(0, str(datetime.now().year + 543))
        LabelChip(row, "อายุ", width=60).pack(side="left", padx=(0, 4))
        self.age_display = ctk.CTkLabel(
            row, text="-", width=60, height=34, fg_color=COLOR_LABEL_CHIP,
            text_color=COLOR_LABEL_TEXT, corner_radius=0, font=ui_font(16, "bold"),
        )
        self.age_display.pack(side="left")

        row = self._row(parent, r); r += 1
        LabelChip(row, "สัญชาติ", width=90).pack(side="left", padx=(0, 4))
        self.entries["nationality"] = make_entry(row, width=90)
        self.entries["nationality"].pack(side="left", padx=(0, 10))
        self.entries["nationality"].insert(0, "ไทย")
        LabelChip(row, "เชื้อชาติ", width=90).pack(side="left", padx=(0, 4))
        self.entries["ethnicity"] = make_entry(row, width=90)
        self.entries["ethnicity"].pack(side="left")
        self.entries["ethnicity"].insert(0, "ไทย")

        self._section_title(parent, r, "ที่อยู่ตามภูมิลำเนาทหาร"); r += 1

        row = self._row(parent, r); r += 1
        LabelChip(row, "บ้านเลขที่", width=90).pack(side="left", padx=(0, 4))
        self.entries["address_no"] = make_entry(row, width=120)
        self.entries["address_no"].pack(side="left", padx=(0, 10))
        LabelChip(row, "หมู่ที่", width=70).pack(side="left", padx=(0, 4))
        self.entries["moo"] = make_entry(row, width=70)
        self.entries["moo"].pack(side="left")

        row = self._row(parent, r); r += 1
        LabelChip(row, "ตำบล", width=90).pack(side="left", padx=(0, 4))
        self.entries["tambon"] = make_entry(row, width=150)
        self.entries["tambon"].pack(side="left", padx=(0, 10))
        LabelChip(row, "อำเภอ", width=90).pack(side="left", padx=(0, 4))
        self.entries["amphoe"] = make_entry(row, width=150)
        self.entries["amphoe"].pack(side="left")

        row = self._row(parent, r); r += 1
        LabelChip(row, "จังหวัด", width=90).pack(side="left", padx=(0, 4))
        self.entries["province"] = make_entry(row, width=150)
        self.entries["province"].pack(side="left", padx=(0, 10))
        LabelChip(row, "เลขบัตร ปชช.", width=110).pack(side="left", padx=(0, 4))
        self.entries["idcard"] = make_entry(row, width=180)
        self.entries["idcard"].pack(side="left")

        self._section_title(parent, r, "บิดา - มารดา"); r += 1

        row = self._row(parent, r); r += 1
        LabelChip(row, "ชื่อบิดา", width=90).pack(side="left", padx=(0, 4))
        self.entries["father_first"] = make_entry(row, width=140)
        self.entries["father_first"].pack(side="left", padx=(0, 10))
        LabelChip(row, "สกุลบิดา", width=90).pack(side="left", padx=(0, 4))
        self.entries["father_last"] = make_entry(row, width=140)
        self.entries["father_last"].pack(side="left")

        row = self._row(parent, r); r += 1
        LabelChip(row, "ชื่อมารดา", width=90).pack(side="left", padx=(0, 4))
        self.entries["mother_first"] = make_entry(row, width=140)
        self.entries["mother_first"].pack(side="left", padx=(0, 10))
        LabelChip(row, "สกุลมารดา", width=90).pack(side="left", padx=(0, 4))
        self.entries["mother_last"] = make_entry(row, width=140)
        self.entries["mother_last"].pack(side="left")

        self._section_title(parent, r, "หมายเลขทหาร"); r += 1

        row = self._row(parent, r); r += 1
        LabelChip(row, "สังกัด,ปี", width=60).pack(side="left", padx=(0, 4))
        self.entries["tb_prefix"] = make_entry(row, width=60)
        self.entries["tb_prefix"].pack(side="left", padx=(0, 6))
        self.entries["tb_prefix"].insert(0, "ท.บ.")
        self.entries["tb_year"] = make_entry(row, width=80)
        self.entries["tb_year"].pack(side="left", padx=(0, 16))
        LabelChip(row, "จังหวัด,ลำดับ", width=70).pack(side="left", padx=(0, 4))
        self.entries["jw_prefix"] = make_entry(row, width=60)
        self.entries["jw_prefix"].pack(side="left", padx=(0, 6))
        self.entries["jw_prefix"].insert(0, "ส.น.")
        self.entries["jw_number"] = make_entry(row, width=80)
        self.entries["jw_number"].pack(side="left")

        self.entries["birth_year"].bind("<KeyRelease>", lambda e: self._update_age())
        self.entries["ref_year"].bind("<KeyRelease>", lambda e: self._update_age())

    # ---------------- TAB 2: ตรวจเลือกทหาร / รด. ----------------
    def _build_conscript_tab(self, tab):
        parent = ctk.CTkScrollableFrame(
            tab, fg_color="transparent",
            scrollbar_button_color=COLOR_BTN_TERTIARY,
            scrollbar_button_hover_color=COLOR_GOLD,
        )
        parent.pack(fill="both", expand=True)
        parent.grid_columnconfigure(0, weight=1)
        r = 0

        self._section_title(parent, r, "ใช้กับเอกสาร: รับรองพลทหาร / รับรอง รด. ปี 3"); r += 1

        row = self._row(parent, r); r += 1
        LabelChip(row, "เกิดวันที่", width=90).pack(side="left", padx=(0, 4))
        self.entries["birth_day"] = make_entry(row, width=55)
        self.entries["birth_day"].pack(side="left", padx=(0, 6))
        self.entries["birth_month"] = make_entry(row, width=110)
        self.entries["birth_month"].pack(side="left", padx=(0, 6))
        self.entries["birth_yearfull"] = make_entry(row, width=80)
        self.entries["birth_yearfull"].pack(side="left")

        row = self._row(parent, r); r += 1
        LabelChip(row, "สังกัด", width=90).pack(side="left", padx=(0, 4))
        self.entries["unit"] = make_entry(row, width=380)
        self.entries["unit"].pack(side="left")

        row = self._row(parent, r); r += 1
        LabelChip(row, "ผลัดที่", width=90).pack(side="left", padx=(0, 4))
        self.entries["conscript_batch"] = make_entry(row, width=80)
        self.entries["conscript_batch"].pack(side="left", padx=(0, 10))
        LabelChip(row, "ประจำปี", width=90).pack(side="left", padx=(0, 4))
        self.entries["conscript_year"] = make_entry(row, width=100)
        self.entries["conscript_year"].pack(side="left")

        row = self._row(parent, r); r += 1
        LabelChip(row, "ปลดวันที่", width=90).pack(side="left", padx=(0, 4))
        self.entries["release_day"] = make_entry(row, width=55)
        self.entries["release_day"].pack(side="left", padx=(0, 6))
        self.entries["release_month"] = make_entry(row, width=110)
        self.entries["release_month"].pack(side="left", padx=(0, 6))
        self.entries["release_year"] = make_entry(row, width=80)
        self.entries["release_year"].pack(side="left")

    # ---------------- TAB 3: ใบแทน สด.43 ----------------
    def _build_sd43_tab(self, tab):
        parent = ctk.CTkScrollableFrame(
            tab, fg_color="transparent",
            scrollbar_button_color=COLOR_BTN_TERTIARY,
            scrollbar_button_hover_color=COLOR_GOLD,
        )
        parent.pack(fill="both", expand=True)
        parent.grid_columnconfigure(0, weight=1)
        r = 0

        self._section_title(parent, r, "ใช้กับเอกสาร: รับรองผลตรวจเลือก (CFRT) / ขอใบแทน สด.43 (NEWRT)"); r += 1

        row = self._row(parent, r); r += 1
        LabelChip(row, "วันที่เกณฑ์", width=100).pack(side="left", padx=(0, 4))
        self.entries["draft_date"] = make_entry(row, width=200)
        self.entries["draft_date"].pack(side="left", padx=(0, 10))
        LabelChip(row, "บัญชี", width=80).pack(side="left", padx=(0, 4))
        self.entries["roll_no"] = make_entry(row, width=340)
        self.entries["roll_no"].pack(side="left")

        row = self._row(parent, r); r += 1
        LabelChip(row, "จำนวนนับที่", width=100).pack(side="left", padx=(0, 4))
        self.entries["count_no"] = make_entry(row, width=100)
        self.entries["count_no"].pack(side="left", padx=(0, 10))
        LabelChip(row, "สด.43 เลขที่", width=100).pack(side="left", padx=(0, 4))
        self.entries["sd43_no"] = make_entry(row, width=120)
        self.entries["sd43_no"].pack(side="left")

        row = self._row(parent, r); r += 1
        LabelChip(row, "คนจำพวกที่", width=100).pack(side="left", padx=(0, 4))
        self.entries["class_type"] = make_entry(row, width=80)
        self.entries["class_type"].pack(side="left", padx=(0, 10))
        LabelChip(row, "ส่วนสูง (ซม.)", width=100).pack(side="left", padx=(0, 4))
        self.entries["height_cm"] = make_entry(row, width=80)
        self.entries["height_cm"].pack(side="left", padx=(0, 10))
        LabelChip(row, "รอบตัว (ซม.)", width=100).pack(side="left", padx=(0, 4))
        self.entries["chest_cm"] = make_entry(row, width=80)
        self.entries["chest_cm"].pack(side="left")

        row = self._row(parent, r); r += 1
        LabelChip(row, "จับสลากได้", width=100).pack(side="left", padx=(0, 4))
        self.entries["lottery_result"] = make_entry(row, width=300)
        self.entries["lottery_result"].pack(side="left")
        ctk.CTkLabel(
            row, text='(กรอก "ดำ" หรือ "แดง")', font=ui_font(13),
            text_color=COLOR_TEXT_MUTED,
        ).pack(side="left", padx=(10, 0))

    def _build_button_bar(self):
        # Export panel อยู่ด้านขวาของหน้าจอ
        panel = ctk.CTkFrame(
            self.workspace, width=380, fg_color=COLOR_SURFACE,
            corner_radius=0, border_width=2, border_color=COLOR_BORDER,
        )
        panel.grid(row=0, column=1, sticky="ns")
        panel.grid_propagate(False)

        # ปักหมุดปุ่ม "เปิดโฟลเดอร์ Export" ไว้ล่างสุดก่อนเป็นอันดับแรก
        # เพื่อไม่ให้ถูกดันหายไปเวลาหน้าต่างเตี้ยลง (แก้ปัญหาปุ่มตก)
        ctk.CTkButton(
            panel, text="เปิดโฟลเดอร์ Export", height=46,
            fg_color=COLOR_BTN_SECONDARY, hover_color=COLOR_BTN_TERTIARY,
            text_color=COLOR_SILVER, corner_radius=0,
            font=ui_font(14, "bold"),
            command=lambda: open_path(OUTPUT_DIR),
        ).pack(side="bottom", fill="x", padx=18, pady=(12, 18))

        ctk.CTkFrame(panel, height=1, fg_color=COLOR_BORDER).pack(
            side="bottom", fill="x", padx=18, pady=(0, 4)
        )

        # ส่วนปุ่ม Export ทั้งหมดอยู่ในกรอบเลื่อนได้ (scrollable) แทนที่จะ
        # ยาวเกินขอบแล้วมองไม่เห็น — ถ้าพื้นที่ไม่พอจะเลื่อนดูได้แทนการหาย
        # หมายเหตุ: กำหนด fg_color ให้ตรงกับพื้นการ์ด (แทน "transparent")
        # และเว้นระยะขอบให้พ้นมุมโค้งของการ์ดด้านนอก (panel corner_radius=0)
        # ป้องกันมุมเหลี่ยมของกรอบเลื่อน (canvas ด้านในเป็นสี่เหลี่ยมเสมอ)
        # โผล่ทะลุออกมานอกมุมโค้งของการ์ด กลายเป็นสี่เหลี่ยมขาวยื่นออกมา
        scroll = ctk.CTkScrollableFrame(
            panel, fg_color=COLOR_SURFACE, corner_radius=0,
            scrollbar_button_color=COLOR_BTN_TERTIARY,
            scrollbar_button_hover_color=COLOR_GOLD,
        )
        scroll.pack(side="top", fill="both", expand=True, padx=10, pady=(10, 0))

        top = ctk.CTkFrame(scroll, fg_color="transparent")
        top.pack(fill="x", padx=18, pady=(18, 10))
        ctk.CTkLabel(
            top, text="EXPORT CENTER", font=pixel_font(13),
            text_color=COLOR_GOLD, anchor="w",
        ).pack(anchor="w")
        ctk.CTkLabel(
            top, text="เลือกเอกสารที่ต้องการสร้าง", font=ui_font(13),
            text_color=COLOR_TEXT_MUTED, anchor="w",
        ).pack(anchor="w", pady=(3, 0))

        self.web_btn = ctk.CTkButton(
            scroll, text="ดึงข้อมูลจากเว็บ TDC", height=44,
            fg_color=COLOR_BTN_SECONDARY, hover_color=COLOR_BTN_TERTIARY,
            text_color=COLOR_GOLD, corner_radius=0, font=ui_font(14, "bold"),
            command=self._pull_from_web,
        )
        self.web_btn.pack(fill="x", padx=18, pady=(0, 8))

        ctk.CTkFrame(scroll, height=1, fg_color=COLOR_BORDER).pack(
            fill="x", padx=18, pady=(4, 12)
        )

        main_docs = [k for k, v in DOCS.items() if v["group"] == "หลัก"]
        sd43_docs = [k for k, v in DOCS.items() if v["group"] == "สด.43"]

        self._export_section(scroll, "เอกสารหลัก", main_docs, primary=True)
        self._export_section(scroll, "เอกสาร สด.43", sd43_docs, primary=False)

    def _export_section(self, parent, title, keys, primary=False):
        box = ctk.CTkFrame(parent, fg_color="transparent")
        box.pack(fill="x", padx=18, pady=(0, 12))
        ctk.CTkLabel(
            box, text=title, font=ui_font(13, "bold"),
            text_color=COLOR_TEXT_MUTED, anchor="w",
        ).pack(anchor="w", pady=(0, 7))

        for key in keys:
            color = COLOR_BTN_PRIMARY if primary else COLOR_BTN_TERTIARY
            hover = COLOR_GOLD_HOVER if primary else "#2E2E66"
            text_color = "#0B0B1A" if primary else COLOR_TEXT_LIGHT
            ctk.CTkButton(
                box,
                text=f"EXPORT\n{DOCS[key]['label']}",
                height=58, corner_radius=0,
                fg_color=color, hover_color=hover, text_color=text_color,
                font=ui_font(13, "bold"),
                anchor="center",
                command=lambda k=key: self._generate(k),
            ).pack(fill="x", pady=4)

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
        self.age_display.configure(text=age if age else "-")

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
            from web_import import WebSession
        except Exception as e:
            messagebox.showerror("ใช้ฟีเจอร์นี้ไม่ได้", f"ขาดไลบรารี playwright หรือโหลดไม่สำเร็จ\n\n{e}\n\n"
                                 "ติดตั้ง: pip install playwright  แล้ว  playwright install chromium")
            return

        sess = getattr(self, "_web", None)
        if sess is not None and sess.is_open and sess.ready.is_set() and not sess.error:
            self._ask_and_fetch()
            return

        # เปิดเบราว์เซอร์ (ครั้งแรกอาจต้องโหลด Chromium 1-3 นาที) แล้วรอจนพร้อมโดยไม่ค้างหน้าจอ
        self._web = WebSession(log=lambda m: self.after(0, lambda: self.status_var.set(m)))
        self._web.open()
        self.web_btn.configure(state="disabled")
        self.status_var.set("กำลังเตรียมเบราว์เซอร์ (ครั้งแรกอาจนาน 1-3 นาที ต้องต่อเน็ต) กรุณารอ...")
        self._wait_web_ready()

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
        messagebox.showinfo(
            "Login เข้า TDC Recruit",
            "ในหน้าต่างเบราว์เซอร์ที่เปิดขึ้น ให้ Login ให้เสร็จก่อน\n"
            "แล้วกด OK ที่นี่ โปรแกรมจะค้นหาและดึงข้อมูลให้ทันที",
        )
        self._ask_and_fetch()

    def _ask_and_fetch(self):
        cid = "".join(ch for ch in self.entries["idcard"].get() if ch.isdigit())
        if len(cid) != 13:
            from tkinter import simpledialog
            cid = simpledialog.askstring("เลขบัตรประชาชน", "กรอกเลขบัตร 13 หลักที่ต้องการดึง:", parent=self) or ""
            cid = "".join(ch for ch in cid if ch.isdigit())
            if len(cid) != 13:
                self.status_var.set("เลขบัตรไม่ครบ 13 หลัก — ยกเลิก")
                return

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
        theme = {
            "bg": COLOR_BG, "surface": COLOR_SURFACE, "text": COLOR_TEXT_LIGHT, "muted": COLOR_TEXT_MUTED,
            "accent": COLOR_GOLD, "accent_hover": COLOR_GOLD_HOVER, "btn2": COLOR_BTN_SECONDARY,
            "btn3": COLOR_BTN_TERTIARY, "entry_bg": COLOR_ENTRY_BG, "entry_border": COLOR_ENTRY_BORDER,
            "entry_text": COLOR_ENTRY_TEXT,
        }
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
    def _generate(self, doc_key: str):
        info = DOCS[doc_key]
        template_path = TEMPLATE_DIR / info["file"]

        if not template_path.exists():
            messagebox.showerror(
                "ไม่พบไฟล์เทมเพลต",
                f"ไม่พบไฟล์ {info['file']} ในโฟลเดอร์:\n{TEMPLATE_DIR}\n\n"
                f"กรุณานำไฟล์ .docx ต้นแบบไปวางในโฟลเดอร์ templates ก่อนใช้งาน",
            )
            return

        if not org_settings.is_configured():
            messagebox.showinfo("ยังไม่ได้ตั้งค่าหน่วยงาน",
                                "กรุณาตั้งค่าหน่วยงาน/ผู้ลงนามก่อนสร้างเอกสาร")
            self._open_settings(first_run=True)
            return

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
            return

        left = self._leftover_placeholders(output_path)
        if left:
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

        if messagebox.askyesno("สำเร็จ", f"สร้างไฟล์ {info['label']} เรียบร้อยแล้ว\nต้องการเปิดไฟล์เลยหรือไม่?"):
            open_path(output_path)

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
