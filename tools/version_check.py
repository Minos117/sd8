"""พิมพ์เลขเวอร์ชันของโปรแกรม และเช็คว่า app.py (APP_VER) ตรงกับ installer.iss (MyAppVersion)
คืนค่า exit code 1 ถ้าไม่ตรง ใช้โดย build_all.bat / MakeRelease"""
import re
import sys
from pathlib import Path

root = Path(__file__).resolve().parent.parent
app = re.search(r'^APP_VER\s*=\s*"([^"]+)"', (root / "app.py").read_text(encoding="utf-8"), re.M)
iss = re.search(r'#define\s+MyAppVersion\s+"([^"]+)"', (root / "installer.iss").read_text(encoding="utf-8-sig"))
if not app or not iss:
    print("ERROR: หา APP_VER ใน app.py หรือ MyAppVersion ใน installer.iss ไม่เจอ", file=sys.stderr)
    sys.exit(1)
if app.group(1) != iss.group(1):
    print(f"ERROR: เวอร์ชันไม่ตรงกัน app.py={app.group(1)} installer.iss={iss.group(1)}", file=sys.stderr)
    sys.exit(1)
print(app.group(1))
