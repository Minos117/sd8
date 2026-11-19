# -*- coding: utf-8 -*-
"""ตรวจ/ดาวน์โหลดอัปเดตจาก GitHub (update.json + SHA-256). แจ้งเท่านั้น ไม่ดาวน์โหลดเอง"""
import hashlib
import json
import os
import tempfile
import urllib.request

DEFAULT_URL = "https://raw.githubusercontent.com/Minos117/sd8/main/update.json"
SUPPORTED_SETUP_REV = 1  # เพิ่มเมื่อโครงสร้าง update.json เปลี่ยนจนโปรแกรมเก่าอ่านไม่ได้


def _vt(v):
    return tuple(int(x) for x in str(v).split("."))


def fetch_info(url=DEFAULT_URL, timeout=10):
    req = urllib.request.Request(url, headers={"User-Agent": "SD8System-Updater"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8-sig"))


def check(app_ver, url=DEFAULT_URL):
    """คืน dict ของ update.json ถ้ามีรุ่นใหม่กว่า, None ถ้าเป็นรุ่นล่าสุด. ผิดพลาด -> raise ข้อความจริง"""
    info = fetch_info(url)
    if int(info.get("setup_rev", 1)) > SUPPORTED_SETUP_REV:
        raise RuntimeError("update.json เป็นรูปแบบใหม่เกินกว่าโปรแกรมนี้อ่านได้ — โหลดตัวติดตั้งใหม่ด้วยมือ")
    return info if _vt(info["version"]) > _vt(app_ver) else None


def download(info, progress=None, timeout=30):
    """โหลดตัวติดตั้ง -> ตรวจ SHA-256 ก่อนเขียนไฟล์จริง. คืน path ของไฟล์ที่ตรวจผ่านแล้ว"""
    url, want = info["url"], str(info["sha256"]).lower()
    if not url:
        raise RuntimeError("update.json ไม่มี url ของตัวติดตั้ง")
    dest = os.path.join(tempfile.gettempdir(), info.get("file", "SD8System_Setup.exe"))
    part = dest + ".part"
    h = hashlib.sha256()
    req = urllib.request.Request(url, headers={"User-Agent": "SD8System-Updater"})
    with urllib.request.urlopen(req, timeout=timeout) as r, open(part, "wb") as f:
        total = int(r.headers.get("Content-Length") or info.get("size") or 0)
        got = 0
        while True:
            b = r.read(1 << 16)
            if not b:
                break
            f.write(b)
            h.update(b)
            got += len(b)
            if progress and total:
                progress(got, total)
    if h.hexdigest() != want:
        os.remove(part)
        raise RuntimeError(f"SHA-256 ไม่ตรง (ได้ {h.hexdigest()[:12]}… ต้องเป็น {want[:12]}…) — ไม่ติดตั้ง")
    if os.path.exists(dest):
        bak = dest + ".bak"
        if os.path.exists(bak):
            os.remove(bak)
        os.replace(dest, bak)
    os.replace(part, dest)
    return dest
