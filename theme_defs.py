"""ชุดธีมของโปรแกรม (10 ธีม) + บันทึก/อ่านธีมที่เลือก
ทุกธีมคือ dict ของสี; app.py นำไปเขียนทับตัวแปร COLOR_* แล้วสร้างหน้าจอใหม่
"""
import json
import os
from pathlib import Path


def _t(label, mode, bg, surface, header, border, chip, chip_text, entry_bg, entry_border,
       text, muted, accent, accent_hover, on_accent, btn2, btn3, btn3_hover, silver,
       primary=None, danger=("#4A1020", "#FF4D6D", "#6A1830"), fonts=(), pixel=False, layout="side"):
    return {
        "label": label, "mode": mode, "pixel": pixel, "layout": layout, "fonts": list(fonts),
        "COLOR_BG": bg, "COLOR_SURFACE": surface, "COLOR_PANEL_HEADER": header,
        "COLOR_CARD": surface, "COLOR_CARD_ALT": btn3, "COLOR_BORDER": border,
        "COLOR_LABEL_CHIP": chip, "COLOR_LABEL_TEXT": chip_text,
        "COLOR_ENTRY_BG": entry_bg, "COLOR_ENTRY_BORDER": entry_border, "COLOR_ENTRY_TEXT": text,
        "COLOR_GOLD": accent, "COLOR_GOLD_HOVER": accent_hover, "COLOR_SILVER": silver,
        "COLOR_BTN_PRIMARY": primary or accent, "COLOR_BTN_SECONDARY": btn2, "COLOR_BTN_TERTIARY": btn3,
        "COLOR_TEXT_LIGHT": text, "COLOR_TEXT_MUTED": muted,
        "COLOR_TAB_SELECTED": accent, "COLOR_TAB_HOVER": accent_hover,
        "COLOR_ON_ACCENT": on_accent, "COLOR_BTN3_HOVER": btn3_hover,
        "COLOR_DANGER_BG": danger[0], "COLOR_DANGER_TEXT": danger[1], "COLOR_DANGER_HOVER": danger[2],
    }


_LIGHT_DANGER = ("#FBE0E0", "#C62828", "#F5C6C6")

THEMES = {
    "retro": _t("เรโทร 8 บิต", "dark", "#0B0B1A", "#14142B", "#0F0F22", "#3A3A8C", "#2B2B6B", "#EDEDFF",
                "#0B0B1A", "#5B5BC8", "#EDEDFF", "#8F8FC0", "#FFD84A", "#FFE98A", "#0B0B1A",
                "#22224A", "#2E2E66", "#3A3A80", "#B8B8E0", fonts=["Chakra Petch"], pixel=True),
    "winxp": _t("Windows XP", "light", "#3A6EA5", "#ECE9D8", "#D6E3F8", "#7F9DB9", "#D4D0C8", "#000000",
                "#FFFFFF", "#7F9DB9", "#000000", "#4D4D4D", "#316AC5", "#4C85DD", "#FFFFFF",
                "#D4D0C8", "#E3DFCF", "#C9C5B2", "#1F3F7A", primary="#3C9A3C",
                danger=_LIGHT_DANGER, fonts=["Tahoma", "Segoe UI"]),
    "psp": _t("PSP (XMB)", "dark", "#06142E", "#0D2347", "#0A1B3A", "#2F5C9E", "#17376B", "#EAF4FF",
              "#081833", "#4F86D6", "#F2F8FF", "#8FB0DC", "#7FD0FF", "#B5E6FF", "#06142E",
              "#12305E", "#183F7A", "#2253A0", "#B8D4F5", fonts=["Segoe UI", "Leelawadee UI"], layout="xmb"),
    "neon": _t("ไซเบอร์เนออน", "dark", "#0A0014", "#150028", "#10001F", "#8A2BE2", "#3A0A5C", "#F5E9FF",
               "#0A0014", "#FF2BD6", "#F5E9FF", "#B58AD9", "#00F0FF", "#7DF8FF", "#0A0014",
               "#26083F", "#3A0F5E", "#5A1A8C", "#D9B8FF", primary="#00F0FF"),
    "terminal": _t("เทอร์มินัลเขียว", "dark", "#000000", "#06120A", "#030A05", "#1F8F3A", "#0C2A14", "#C8FFD4",
                   "#000000", "#1F8F3A", "#C8FFD4", "#4FBF6A", "#39FF14", "#8CFF73", "#000000",
                   "#0C2A14", "#103A1C", "#175A2A", "#7FE08F"),
    "ocean": _t("มหาสมุทร (สว่าง)", "light", "#E8F1FA", "#FFFFFF", "#DCEBFA", "#9CC2E5", "#D2E6F9", "#10243A",
                "#FFFFFF", "#7DAEDD", "#10243A", "#5A7390", "#0B6FD4", "#3C8DE0", "#FFFFFF",
                "#E1EEFA", "#EAF3FC", "#D2E6F9", "#3F5F80", danger=_LIGHT_DANGER),
    "sakura": _t("ซากุระ (สว่าง)", "light", "#FDF1F5", "#FFFFFF", "#FBE3EB", "#F0A6BE", "#FADDE6", "#3A1626",
                 "#FFFFFF", "#E58AA8", "#3A1626", "#8C5A6E", "#D6336C", "#E8588A", "#FFFFFF",
                 "#FBE3EB", "#FDEEF3", "#FADDE6", "#8C2F52", danger=_LIGHT_DANGER),
    "sunset": _t("พระอาทิตย์ตก", "dark", "#1B0E0A", "#2A1610", "#22120D", "#C2551E", "#4A2214", "#FFE9D6",
                 "#1B0E0A", "#E07A2F", "#FFE9D6", "#C9987A", "#FF9F1C", "#FFBE5C", "#1B0E0A",
                 "#3A1D12", "#4A2418", "#6A3420", "#F0C8A8"),
    "forest": _t("ป่าลึก", "dark", "#0B1A12", "#12261A", "#0E2016", "#2E7D4F", "#1D4630", "#E6F5E9",
                 "#0B1A12", "#3FA066", "#E6F5E9", "#86B494", "#9BE564", "#BDF08F", "#0B1A12",
                 "#183324", "#1F4430", "#2A5C40", "#B5DDC0"),
    "royal": _t("รอยัลโกลด์", "dark", "#0D0D0D", "#171717", "#121212", "#8C6D1F", "#2A2310", "#F5EFD8",
                "#0D0D0D", "#B8922C", "#F5EFD8", "#A8997A", "#D4AF37", "#F0D060", "#0D0D0D",
                "#242424", "#2F2A18", "#44391C", "#CDBE95"),
}
for _k, _c in {"retro": 0, "terminal": 0, "winxp": 3, "psp": 10, "neon": 4, "ocean": 8,
               "sakura": 12, "sunset": 8, "forest": 8, "royal": 4}.items():
    THEMES[_k]["corner"] = _c   # ความมนของมุม (px)
# ความหรู/สมจริงเพิ่ม: แถบหัวไล่สี, หัวข้อกลุ่มแบบ Explorer, ขอบปุ่ม, ความหนาขอบช่องกรอก
THEMES["winxp"].update(
    hdr_grad=("#0A3FC8", "#3F8EF5"), hdr_text="#FFFFFF", hdr_sub="#DCE8FF",
    group_grad=("#FFFFFF", "#C6D3F7"), group_text="#215DC6", btn_border="#003C74", entry_bw=1)
THEMES["psp"].update(hdr_grad=("#050F24", "#1F5AA8"), hdr_text="#FFFFFF", hdr_sub="#B5D4FF")
THEMES["neon"].update(hdr_grad=("#14002A", "#5A0B8C"), hdr_text="#00F0FF", hdr_sub="#F5E9FF")
DEFAULT_THEME = "retro"
ORDER = ["retro", "winxp", "psp", "neon", "terminal", "ocean", "sakura", "sunset", "forest", "royal"]


def _path(base: Path) -> Path:
    return Path(base) / "theme.json"


def load_theme(dirs) -> str:
    for d in dirs:
        try:
            p = _path(d)
            if p.exists():
                name = json.loads(p.read_text(encoding="utf-8")).get("theme", "")
                if name in THEMES:
                    return name
        except Exception:
            pass
    return DEFAULT_THEME


def save_theme(dirs, name: str) -> bool:
    for d in dirs:
        try:
            Path(d).mkdir(parents=True, exist_ok=True)
            tmp = _path(d).with_suffix(".tmp")
            tmp.write_text(json.dumps({"theme": name}, ensure_ascii=False), encoding="utf-8")
            os.replace(tmp, _path(d))
            return True
        except Exception:
            continue
    return False
