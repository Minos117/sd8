# -*- coding: utf-8 -*-
"""ตรวจ/เปิดโปรแกรม Linkage2Client (Windows) ก่อนดึงข้อมูลจาก Recruit
ไม่รู้ชื่อโปรเซส/ตำแหน่งติดตั้งแน่ชัด -> ค้นหาตามชื่อที่มีคำว่า linkage แล้วจำ path ไว้ (ถามครั้งเดียว)"""
import glob
import os
import platform
import subprocess
import time
from pathlib import Path

KEY = "linkage"


def _store() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    d = Path(base) / "SD8System"
    d.mkdir(parents=True, exist_ok=True)
    return d / "linkage_path.txt"


def saved_path() -> str:
    try:
        p = _store().read_text(encoding="utf-8-sig").strip()
        return p if p and os.path.isfile(p) else ""
    except Exception:
        return ""


def save_path(p: str):
    _store().write_text(p, encoding="utf-8")


def running_names() -> list:
    """ชื่อโปรเซสที่มีคำว่า linkage ที่กำลังทำงาน (ว่าง = ไม่ได้เปิด)"""
    if platform.system() != "Windows":
        return []
    try:
        out = subprocess.run(["tasklist", "/FO", "CSV", "/NH"], capture_output=True, timeout=15,
                             creationflags=0x08000000).stdout.decode("cp874", "ignore")
    except Exception:
        return []
    names = {line.split('","')[0].strip('"') for line in out.splitlines() if line.strip()}
    return sorted(n for n in names if KEY in n.lower())


def find_exe() -> str:
    p = saved_path()
    if p:
        return p
    la = os.environ.get("LOCALAPPDATA")
    if la:
        known = os.path.join(la, "Programs", "Linkage2Client", "Linkage2Client.exe")
        if os.path.isfile(known):
            return known
    roots = [os.environ.get(v) for v in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA", "ProgramData")]
    roots += ["C:\\"]
    for r in filter(None, roots):
        for pat in (f"*{KEY}*\\*.exe", f"*{KEY}*\\**\\*.exe", f"Programs\\*{KEY}*\\*.exe"):
            for f in glob.glob(os.path.join(r, pat), recursive=True):
                n = os.path.basename(f).lower()
                if "unins" not in n and "setup" not in n:
                    return f
    return ""


def ensure_running(wait=12):
    """คืน (สถานะ, ข้อความ): 'running' | 'started' | 'notfound' | 'failed'"""
    if platform.system() != "Windows":
        return "failed", "ใช้ได้เฉพาะ Windows"
    if running_names():
        return "running", "Linkage2Client เปิดอยู่แล้ว"
    exe = find_exe()
    if not exe:
        return "notfound", "หาตัวโปรแกรม Linkage2Client ที่ติดตั้งไว้ไม่เจอ"
    try:
        subprocess.Popen([exe], cwd=os.path.dirname(exe))
    except Exception as e:
        return "failed", f"เปิด {exe} ไม่สำเร็จ: {e}"
    for _ in range(wait):
        time.sleep(1)
        if running_names():
            return "started", f"เปิด Linkage2Client ให้แล้ว ({exe})"
    return "started", f"สั่งเปิดแล้ว แต่ยังไม่เห็นโปรเซสชื่อ linkage ({exe}) — อาจใช้ชื่ออื่น"


def port() -> int:
    """พอร์ตของ Linkage2Client จาก appsettings.json ข้างตัว exe (ไม่เจอ = 15043)"""
    import json
    exe = find_exe()
    if exe:
        try:
            p = os.path.join(os.path.dirname(exe), "appsettings.json")
            return int(json.loads(open(p, encoding="utf-8-sig").read())["AppConfig"]["Port"])
        except Exception:
            pass
    return 15043
