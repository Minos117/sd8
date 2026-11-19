"""settings_dialog.py - หน้าต่างตั้งค่าหน่วยงาน/ผู้ลงนาม (เปิดอัตโนมัติเมื่อใช้งานครั้งแรก)"""

from tkinter import messagebox

import customtkinter as ctk

import org_settings as org


class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent, theme, font, current, on_saved, first_run=False, on_cancel=None):
        super().__init__(parent)
        self.theme, self.font = theme, font
        self.on_saved, self.on_cancel = on_saved, on_cancel
        self.first_run = first_run
        self.entries = {}
        self._last_auto = org.derive(current)

        self.title("ตั้งค่าหน่วยงาน" + (" — ใช้งานครั้งแรก" if first_run else ""))
        # ขนาดหน้าต่างปรับตามจอ (หน้าจอเล็ก/ขยาย DPI) กันปุ่มล่างตกขอบจอ
        try:
            scale = float(self._get_window_scaling())
        except Exception:
            scale = 1.0
        scale = scale if scale > 0 else 1.0
        sw, sh = self.winfo_screenwidth() / scale, self.winfo_screenheight() / scale
        w, h = int(min(780, sw * 0.92)), int(min(820, sh * 0.80))   # เผื่อแถบงาน Windows ~ 20%
        self.geometry(f"{w}x{h}+{int(max(0, (sw - w) / 2 * scale))}+{int(max(0, (sh - h) / 2 * scale * 0.6))}")
        self.minsize(min(640, w), min(420, h))
        self.configure(fg_color=theme["bg"])
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        head = ctk.CTkFrame(self, fg_color="transparent")
        head.pack(fill="x", padx=20, pady=(16, 6))
        ctk.CTkLabel(head, text="ตั้งค่าหน่วยงาน", font=font(22, "bold"),
                     text_color=theme["text"], anchor="w").pack(anchor="w")
        tip = ("กรอก อำเภอ / จังหวัด / ตัวย่อจังหวัด ก่อน ช่องอื่นจะเติมให้อัตโนมัติ (แก้ทับได้) "
               "ค่านี้จะถูกบันทึกไว้ ครั้งต่อไปไม่ต้องกรอกใหม่ และแก้ได้จากปุ่ม \"ตั้งค่าหน่วยงาน\"")
        ctk.CTkLabel(head, text=tip, font=font(13), text_color=theme["muted"],
                     anchor="w", justify="left", wraplength=max(360, w - 70)).pack(anchor="w", pady=(2, 0))

        # ปุ่มด้านล่างปักก่อน กันถูกดันหาย
        foot = ctk.CTkFrame(self, fg_color="transparent")
        foot.pack(side="bottom", fill="x", padx=20, pady=(6, 16))
        self.date_lbl = ctk.CTkLabel(foot, text="", font=font(12), text_color=theme["muted"], anchor="w")
        self.date_lbl.pack(fill="x", pady=(0, 8))
        ctk.CTkButton(foot, text="บันทึก", height=42, width=150, font=font(15, "bold"),
                      fg_color=theme["accent"], hover_color=theme["accent_hover"], corner_radius=0,
                      text_color="#0B0B1A", command=self._save).pack(side="right")
        ctk.CTkButton(foot, text="ยกเลิก", height=42, width=100, font=font(14, "bold"),
                      fg_color=theme["btn2"], hover_color=theme["btn3"], text_color=theme["text"], corner_radius=0,
                      command=self._cancel).pack(side="right", padx=(0, 8))
        ctk.CTkButton(foot, text="โหลดค่าตัวอย่าง (วานรนิวาส)", height=42, font=font(13),
                      fg_color=theme["btn2"], hover_color=theme["btn3"], text_color=theme["muted"], corner_radius=0,
                      command=self._load_sample).pack(side="left")

        body = ctk.CTkScrollableFrame(self, fg_color=theme["surface"], corner_radius=0)
        body.pack(fill="both", expand=True, padx=20, pady=6)
        body.grid_columnconfigure(1, weight=1)

        r = 0
        for title, items in org.GROUPS:
            ctk.CTkLabel(body, text=title, font=font(15, "bold"), text_color=theme["accent"],
                         anchor="w").grid(row=r, column=0, columnspan=2, sticky="w", padx=12, pady=(14, 4))
            r += 1
            for key, label, hint, required in items:
                ctk.CTkLabel(body, text=label + (" *" if required else ""), font=font(13),
                             text_color=theme["text"], anchor="w").grid(
                    row=r, column=0, sticky="w", padx=(12, 8), pady=3)
                e = ctk.CTkEntry(body, placeholder_text="เช่น " + hint, font=font(14), height=34,
                                 fg_color=theme["entry_bg"], border_color=theme["entry_border"], border_width=2, corner_radius=0,
                                 text_color=theme["entry_text"])
                e.grid(row=r, column=1, sticky="we", padx=(0, 12), pady=3)
                if current.get(key):
                    e.insert(0, current[key])
                self.entries[key] = e
                if key in ("amphoe", "province", "prov_abbr"):
                    e.bind("<KeyRelease>", lambda _e: self._refresh_derived())
                r += 1
        ctk.CTkLabel(body, text="* = ต้องกรอก", font=font(12), text_color=theme["muted"],
                     anchor="w").grid(row=r, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 8))

        self._update_date_preview()
        self.after(150, self._grab)

    # ---------------------------------------------------------------
    def _grab(self):
        try:
            self.grab_set()
            self.focus_force()
            self.lift()
        except Exception:
            pass

    def _values(self):
        return {k: e.get().strip() for k, e in self.entries.items()}

    def _set(self, key, text):
        e = self.entries[key]
        e.delete(0, "end")
        if text:
            e.insert(0, text)

    def _refresh_derived(self):
        new = org.derive(self._values())
        for k, nv in new.items():
            cur = self.entries[k].get().strip()
            if cur == "" or cur == self._last_auto.get(k):
                self._set(k, nv)
                self._last_auto[k] = nv

    def _update_date_preview(self):
        d = org.date_mapping()
        self.date_lbl.configure(
            text=f"วันที่ที่ระบบใส่ให้อัตโนมัติ (ไม่มีวัน): {d['เดือนย่อปีย่อ']}  |  {d['เดือนเต็มปีเต็ม']}  |  /{d['ปีอารบิก']}")

    def _load_sample(self):
        if any(self._values().values()) and not messagebox.askyesno(
                "โหลดค่าตัวอย่าง", "ช่องที่กรอกไว้จะถูกเขียนทับด้วยค่าตัวอย่าง ต้องการต่อหรือไม่?", parent=self):
            return
        for k, v in org.SAMPLE.items():
            self._set(k, v)
        self._last_auto = org.derive(org.SAMPLE)

    def _save(self):
        vals = self._values()
        miss = org.missing_required(vals)
        if miss:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอก:\n- " + "\n- ".join(miss), parent=self)
            return
        try:
            path = org.save(vals)
        except Exception as e:
            messagebox.showerror("บันทึกไม่สำเร็จ", str(e), parent=self)
            return
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()
        self.on_saved(vals, path)

    def _cancel(self):
        if self.first_run and not messagebox.askyesno(
                "ยังไม่ได้ตั้งค่า",
                "ถ้ายังไม่ตั้งค่า จะสร้างเอกสารไม่ได้\nต้องการปิดหน้านี้ไว้ก่อนใช่หรือไม่?", parent=self):
            return
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()
        if self.on_cancel:
            self.on_cancel()


def thai_id_valid(digits: str) -> bool:
    """ตรวจเลขบัตรประชาชน 13 หลักด้วยหลักตรวจสอบมาตรฐาน"""
    if len(digits) != 13 or not digits.isdigit():
        return False
    s = sum(int(digits[i]) * (13 - i) for i in range(12))
    return (11 - s % 11) % 10 == int(digits[12])


def _fmt_id(d: str) -> str:
    """ขีดคั่นแทรกเฉพาะเมื่อมีเลขตัวถัดไปแล้ว (ไม่มีขีดค้างท้าย): 1-2345-67890-12-3"""
    parts, last = [], 0
    for c in (1, 5, 10, 12, 13):
        if len(d) > last:
            parts.append(d[last:c])
        last = c
    return "-".join(parts)


class IdPromptDialog(ctk.CTkToplevel):
    """ถามเลขบัตรประชาชนก่อนดึงข้อมูลจากเว็บ: พิมพ์/วางได้ จัดรูปแบบอัตโนมัติ ตรวจหลักตรวจสอบ Enter=ดึง Esc=ยกเลิก"""

    def __init__(self, parent, theme, font, initial, on_ok, on_cancel=None):
        super().__init__(parent)
        self.on_ok, self.on_cancel = on_ok, on_cancel
        self.var = ctk.StringVar()
        self._busy = False
        self.title("ดึงข้อมูลจากเว็บ TDC")
        self.configure(fg_color=theme["bg"])
        self.resizable(False, False)
        self.transient(parent)
        w, h = 520, 270
        try:
            scale = float(self._get_window_scaling()) or 1.0
        except Exception:
            scale = 1.0
        px = parent.winfo_rootx() + max(0, (parent.winfo_width() - int(w * scale)) // 2)
        py = parent.winfo_rooty() + max(0, (parent.winfo_height() - int(h * scale)) // 3)
        self.geometry(f"{w}x{h}+{px}+{py}")
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        box = ctk.CTkFrame(self, fg_color=theme["surface"], corner_radius=0,
                           border_width=2, border_color=theme["entry_border"])
        box.pack(fill="both", expand=True, padx=14, pady=14)
        ctk.CTkLabel(box, text="เลขบัตรประชาชนที่ต้องการดึง", font=font(18, "bold"),
                     text_color=theme["accent"], anchor="w").pack(anchor="w", padx=18, pady=(16, 2))
        ctk.CTkLabel(box, text="พิมพ์หรือวาง (Ctrl+V) 13 หลัก ขีดคั่นใส่หรือไม่ใส่ก็ได้", font=font(13),
                     text_color=theme["muted"], anchor="w").pack(anchor="w", padx=18)
        self.entry = ctk.CTkEntry(box, textvariable=self.var, height=46, width=420, corner_radius=0,
                                  fg_color=theme["entry_bg"], text_color=theme["entry_text"],
                                  border_width=2, border_color=theme["entry_border"], font=font(22, "bold"))
        self.entry.pack(padx=18, pady=(12, 4))
        self.hint = ctk.CTkLabel(box, text="", font=font(13), text_color=theme["muted"], anchor="w")
        self.hint.pack(anchor="w", padx=18)

        row = ctk.CTkFrame(box, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=(10, 14), side="bottom")
        self.ok_btn = ctk.CTkButton(row, text="ดึงข้อมูล", height=42, width=150, corner_radius=0,
                                    fg_color=theme["accent"], hover_color=theme["accent_hover"],
                                    text_color="#0B0B1A", font=font(15, "bold"), command=self._ok)
        self.ok_btn.pack(side="right")
        ctk.CTkButton(row, text="ยกเลิก", height=42, width=100, corner_radius=0,
                      fg_color=theme["btn2"], hover_color=theme["btn3"], text_color=theme["text"],
                      font=font(14, "bold"), command=self._cancel).pack(side="right", padx=(0, 8))

        self.var.trace_add("write", self._on_change)
        self.var.set("".join(c for c in (initial or "") if c.isdigit())[:13])
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self._cancel())
        self.after(80, self._focus)

    def _focus(self):
        try:
            self.grab_set()
            self.entry.focus_force()
            self.entry.icursor("end")
        except Exception:
            pass

    def digits(self):
        return "".join(c for c in self.var.get() if c.isdigit())

    def _on_change(self, *_):
        if self._busy:
            return
        d = self.digits()[:13]
        shown = _fmt_id(d)
        if shown != self.var.get():
            self._busy = True
            self.var.set(shown)
            self._busy = False
            try:
                self.entry.icursor("end")
                self.after_idle(lambda: self.entry.icursor("end"))
            except Exception:
                pass
        if len(d) < 13:
            self.hint.configure(text=f"{len(d)}/13 หลัก", text_color="#9AA0C8")
            self.ok_btn.configure(state="disabled")
        elif not thai_id_valid(d):
            self.hint.configure(text="เลขนี้ไม่ผ่านการตรวจหลักตรวจสอบ — เช็กตัวเลขอีกครั้ง", text_color="#FF4D6D")
            self.ok_btn.configure(state="disabled")
        else:
            self.hint.configure(text="เลขถูกต้อง ✓ กด Enter เพื่อดึงข้อมูล", text_color="#5DFF9A")
            self.ok_btn.configure(state="normal")

    def _ok(self):
        d = self.digits()
        if not thai_id_valid(d):
            return
        self.grab_release()
        self.destroy()
        self.on_ok(d)

    def _cancel(self):
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()
        if self.on_cancel:
            self.on_cancel()
