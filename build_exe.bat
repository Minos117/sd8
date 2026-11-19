@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo   Build standalone EXE (PyInstaller)
echo ============================================
echo.

REM --------------------------------------------
REM 1) ตรวจสอบว่ามี Python ให้ใช้งานจริง
REM --------------------------------------------
where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] ไม่พบ python ใน PATH
    echo กรุณาติดตั้ง Python จาก https://www.python.org/downloads/
    echo และตอนติดตั้งให้ติ๊ก "Add python.exe to PATH" ด้วย
    pause
    exit /b 1
)

REM --------------------------------------------
REM 2) สร้าง virtual environment (ถ้ายังไม่มี)
REM --------------------------------------------
if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERROR] สร้าง virtual environment ไม่สำเร็จ
        pause
        exit /b 1
    )
)

if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] ไม่พบ .venv\Scripts\activate.bat
    echo virtual environment อาจสร้างไม่สมบูรณ์ ลองลบโฟลเดอร์ .venv แล้วรันใหม่
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo [ERROR] เปิดใช้งาน virtual environment ไม่สำเร็จ
    pause
    exit /b 1
)

REM --------------------------------------------
REM 3) อัปเดต pip ก่อน (ป้องกันปัญหา install ล้มเหลว
REM    จาก pip เวอร์ชันเก่าเกินไป)
REM --------------------------------------------
echo.
echo Upgrading pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo [WARNING] อัปเกรด pip ไม่สำเร็จ จะลองติดตั้ง dependencies ต่อไป
)

REM --------------------------------------------
REM 4) ติดตั้ง dependencies — ต้องเช็คผลลัพธ์ให้ชัดเจน
REM    (ก่อนหน้านี้ ถ้าขั้นตอนนี้ล้มเหลว สคริปต์จะไม่หยุด
REM    และไปพังตอน PyInstaller โดยไม่รู้สาเหตุจริง)
REM --------------------------------------------
echo.
echo Installing/updating dependencies...
echo.

if exist "requirements.txt" (
    python -m pip install -r requirements.txt
) else (
    python -m pip install customtkinter pillow python-docx openpyxl lxml
)

if errorlevel 1 (
    echo.
    echo [ERROR] ติดตั้ง dependencies ไม่สำเร็จ ^(pip install ล้มเหลว^)
    echo สาเหตุที่พบบ่อย:
    echo   - ไม่มีอินเทอร์เน็ต หรือถูกไฟร์วอลล์/พร็อกซีบล็อก
    echo   - เวอร์ชัน Python ใหม่/เก่าเกินไปจนไม่มี wheel ให้แพ็กเกจบางตัว
    echo   - เขียนไฟล์ใน .venv ไม่ได้ ^(สิทธิ์ไม่พอ / โฟลเดอร์ถูกล็อก^)
    echo ลองรัน "python -m pip install customtkinter" มือเปล่าดู error เต็มๆ
    pause
    exit /b 1
)

REM --------------------------------------------
REM 4b) ติดตั้ง pyinstaller แยกต่างหากเสมอ ไม่ว่า requirements.txt
REM     จะมีมันอยู่หรือไม่ — เพราะ pyinstaller เป็นเครื่องมือ build
REM     ไม่ใช่ dependency ของตัวแอป ถ้า requirements.txt ไม่มีรายการนี้
REM     (พบบ่อยมาก) การ build จะพังทันทีที่ขั้นตอนถัดไปโดยไม่รู้สาเหตุ
REM --------------------------------------------
echo.
echo Ensuring PyInstaller is installed...
python -m pip install pyinstaller
if errorlevel 1 (
    echo.
    echo [ERROR] ติดตั้ง pyinstaller ไม่สำเร็จ
    pause
    exit /b 1
)

REM --------------------------------------------
REM 4c) ตรวจสอบซ้ำว่าแพ็กเกจที่แอปต้องใช้จริง import ได้ครบ
REM     (เผื่อ requirements.txt เก่า/ไม่ครบ แต่ pip install ผ่านแบบไม่ error
REM     เพราะแค่ไม่มีรายการนั้นอยู่เลย ไม่ใช่ install แล้ว fail)
REM --------------------------------------------
echo.
echo Verifying required packages...
python -c "import customtkinter" 2>nul || (echo Installing customtkinter... & python -m pip install customtkinter)
python -c "import PIL" 2>nul || (echo Installing pillow... & python -m pip install pillow)
python -c "import docx" 2>nul || (echo Installing python-docx... & python -m pip install python-docx)
python -c "import openpyxl" 2>nul || (echo Installing openpyxl... & python -m pip install openpyxl)
python -c "import lxml" 2>nul || (echo Installing lxml... & python -m pip install lxml)

REM --------------------------------------------
REM 5) ตรวจสอบว่า pyinstaller ใช้งานได้จริงก่อนเริ่ม build
REM --------------------------------------------
python -m PyInstaller --version >nul 2>nul
if errorlevel 1 (
    echo [ERROR] ติดตั้ง PyInstaller แล้วแต่เรียกใช้งานไม่ได้
    pause
    exit /b 1
)

echo.
echo Cleaning old build folders...
echo.

REM --------------------------------------------
REM 4d) ลบโฟลเดอร์ build/dist เอง แบบ retry เพราะถ้าโปรเจกต์นี้
REM     อยู่ใน OneDrive (หรือมี antivirus/Explorer ล็อกไฟล์ค้างอยู่)
REM     คำสั่ง rmdir อาจล้มเหลวชั่วคราวด้วย "Access is denied"
REM     เราจึงลองซ้ำหลายครั้งก่อนจะยอมแพ้ แทนที่จะปล่อยให้ PyInstaller
REM     ไปพังตอน --clean (ซึ่งจะไม่ retry และทำให้ build ล้มเหลวทันที)
REM --------------------------------------------
call :safe_rmdir "build"
call :safe_rmdir "dist"

python -m PyInstaller --noconfirm --onedir --windowed ^
    --name "SD8System" ^
    --icon "favicon.ico" ^
    --collect-all customtkinter ^
    --collect-all PIL ^
    --hidden-import openpyxl ^
    --hidden-import lxml ^
    --hidden-import docx ^
    --hidden-import docx_engine ^
    --hidden-import utils ^
    --hidden-import storage ^
    app.py

if not exist "dist\SD8System\SD8System.exe" (
    echo.
    echo [ERROR] BUILD FAILED! ดู error ด้านบนเพื่อหาสาเหตุ
    pause
    exit /b 1
)

echo.
echo Copying templates...
if exist "templates" (
    xcopy /E /I /Y templates "dist\SD8System\templates" >nul
) else (
    echo [WARNING] ไม่พบโฟลเดอร์ templates — ต้องนำไฟล์ .docx ต้นแบบไปวางเองใน dist\SD8System\templates
)

if not exist "dist\SD8System\output" mkdir "dist\SD8System\output"

echo.
echo ============================================
echo   BUILD SUCCESS!
echo   dist\SD8System\SD8System.exe
echo.
echo   หมายเหตุ: ถ้าเปิด .exe แล้วหน้าต่างไม่ขึ้น/ปิดเงียบๆ
echo   ให้เช็คไฟล์ dist\SD8System\crash_log.txt (ถ้ามี)
echo   หรือใช้ build_exe_debug.bat เพื่อ build แบบมี console
echo   จะได้เห็น error เต็มๆ ตรงๆ
echo ============================================
pause
exit /b 0

REM ==============================================
REM Subroutine: safe_rmdir <folder>
REM ลบโฟลเดอร์แบบ retry สูงสุด 5 ครั้ง ห่างกันครั้งละ ~2 วินาที
REM เพื่อรองรับกรณีไฟล์ถูกล็อกชั่วคราวโดย OneDrive / antivirus /
REM Explorer ที่เปิดดูโฟลเดอร์ค้างอยู่
REM ==============================================
:safe_rmdir
set "_target=%~1"
if not exist "%_target%" goto :eof
set "_tries=0"
:safe_rmdir_loop
rmdir /s /q "%_target%" 2>nul
if not exist "%_target%" goto :eof
set /a _tries+=1
if %_tries% geq 5 (
    echo.
    echo [ERROR] ลบโฟลเดอร์ "%_target%" ไม่สำเร็จหลังจากลอง 5 ครั้ง
    echo สาเหตุที่พบบ่อยที่สุด: โปรเจกต์นี้อยู่ในโฟลเดอร์ OneDrive
    echo ที่กำลัง sync อยู่ ทำให้ไฟล์ถูกล็อกชั่วคราว
    echo แนวทางแก้:
    echo   1^) ย้ายโปรเจกต์ออกจาก OneDrive ไปไว้ที่อื่น เช่น C:\Dev\...
    echo   2^) หรือ Pause syncing ใน OneDrive ชั่วคราวระหว่าง build
    echo   3^) ปิดหน้าต่าง Explorer ที่เปิดค้างอยู่ในโฟลเดอร์ "%_target%"
    echo   4^) เพิ่ม exclusion ให้ antivirus ไม่สแกนโฟลเดอร์โปรเจกต์นี้
    pause
    exit /b 1
)
echo   ลบ "%_target%" ไม่สำเร็จ ลองใหม่ ^(%_tries%/5^)...
timeout /t 2 /nobreak >nul
goto :safe_rmdir_loop
