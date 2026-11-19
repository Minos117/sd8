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
