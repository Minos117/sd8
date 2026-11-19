วิธี build (Windows)
=====================
1. ติดตั้ง Python 3.11-3.13 (ติ๊ก Add python.exe to PATH)
2. (ถ้าต้องการตัวติดตั้ง) ติดตั้ง Inno Setup 6 จาก https://jrsoftware.org/isdl.php
3. ดับเบิลคลิก build_all.bat
   ผลลัพธ์:
     dist\SD8System\SD8System.exe        โปรแกรม (โฟลเดอร์เดียว ใช้ได้เลย)
     Output\SD8System_Setup.exe          ตัวติดตั้ง
     release\SD8System_Beta_<เวอร์ชัน>.zip   ตัวติดตั้ง + update.json (มี SHA-256)

เลขเวอร์ชัน: ต้องตรงกันทั้ง app.py (APP_VER) และ installer.iss (MyAppVersion)
รุ่นแรกเป็น Beta 0.9.x - รุ่น Stable 1.0.0 เมื่อสั่งเท่านั้น

ใช้ครั้งแรก
-----------
- เปิดโปรแกรม -> หน้า "ตั้งค่าหน่วยงาน" เด้งขึ้นเอง กรอกครั้งเดียว (แก้ภายหลังได้จากปุ่มบนหัวโปรแกรม)
- ปุ่ม "ดึงข้อมูลจากเว็บ TDC" ใช้ Playwright/Chromium ครั้งแรกจะดาวน์โหลด Chromium (ต้องต่อเน็ต)

ถ้า Windows Defender / แอนตี้ไวรัสบล็อก .exe (false positive)
-------------------------------------------------------------
- โปรแกรมที่แพ็กด้วย PyInstaller มักถูกเตือนผิดพลาดได้ ไม่ได้แปลว่ามีไวรัส
- รุ่นนี้ปิด UPX แล้ว (ลดโอกาสโดนเตือน) และใช้แบบโฟลเดอร์ (onedir) ไม่ใช่ onefile
- ถ้ายังโดน: ส่งไฟล์ให้ Microsoft ที่ https://www.microsoft.com/wdsi/filesubmission
  เลือก "Software developer" แจ้งเป็น false positive หรือเพิ่มโฟลเดอร์โปรแกรมใน Exclusions ของเครื่องที่ใช้
- ตรวจ SHA-256 ของตัวติดตั้งให้ตรงกับใน update.json ก่อนแจกจ่าย

ฟอนต์ (OFL): Chakra Petch, Press Start 2P อยู่ในโฟลเดอร์ fonts (ใช้เฉพาะในโปรแกรมนี้)
