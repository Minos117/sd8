@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Build ALL - SD8System

echo ============================================
echo   BUILD ALL : EXE + Installer + Release ZIP
echo ============================================
echo.

REM ---------- 1) Python ----------
where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] ไม่พบ python ใน PATH
    echo ติดตั้งจาก https://www.python.org/downloads/ และติ๊ก "Add python.exe to PATH"
    goto :fail
)

REM ---------- 2) เวอร์ชันต้องตรงกัน app.py / installer.iss ----------
set "APP_VER="
for /f "delims=" %%v in ('python tools\version_check.py') do set "APP_VER=%%v"
if not defined APP_VER (
    echo [ERROR] เวอร์ชันใน app.py กับ installer.iss ไม่ตรงกัน หรืออ่านไม่ได้
    python tools\version_check.py
    goto :fail
)
echo เวอร์ชัน: Beta %APP_VER%
echo.

REM ---------- 3) virtual environment ----------
if not exist ".venv\Scripts\python.exe" (
    echo สร้าง virtual environment...
    python -m venv .venv
    if errorlevel 1 goto :fail
)
call .venv\Scripts\activate.bat
if errorlevel 1 goto :fail

REM ---------- 4) ไลบรารี ----------
echo.
echo ติดตั้ง/อัปเดตไลบรารี...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] ติดตั้งไลบรารีไม่สำเร็จ - เช็คอินเทอร์เน็ต/พร็อกซี
    goto :fail
)
python -m pip install pyinstaller
if errorlevel 1 goto :fail

REM ---------- 5) ตรวจโค้ดก่อน build ----------
echo.
echo ตรวจ syntax...
python -m py_compile app.py web_import.py org_settings.py settings_dialog.py docx_engine.py utils.py storage.py
if errorlevel 1 (
    echo [ERROR] โค้ดมี syntax error - แก้ก่อน build
    goto :fail
)

REM ---------- 6) ล้างของเก่า ----------
echo.
echo ล้าง build/dist เก่า...
call :safe_rmdir "build"
call :safe_rmdir "dist"

REM ---------- 7) PyInstaller (ใช้ SD8System.spec) ----------
echo.
echo สร้าง EXE ด้วย PyInstaller...
python -m PyInstaller --noconfirm --clean SD8System.spec
if not exist "dist\SD8System\SD8System.exe" (
    echo [ERROR] BUILD FAILED - ดู error ด้านบน
    goto :fail
)

REM ---------- 8) คัดลอกไฟล์ที่โปรแกรมอ่านจากข้าง exe ----------
echo.
echo คัดลอก templates / fonts / logo...
if not exist "templates\SD8.docx" (
    echo [ERROR] ไม่พบ templates\*.docx
    goto :fail
)
xcopy /E /I /Y templates "dist\SD8System\templates" >nul
if exist "fonts" xcopy /E /I /Y fonts "dist\SD8System\fonts" >nul
copy /Y sd8_logo.png "dist\SD8System\" >nul
copy /Y favicon.ico "dist\SD8System\" >nul
if not exist "dist\SD8System\output" mkdir "dist\SD8System\output"

REM ---------- 9) Inno Setup ----------
echo.
set "ISCC="
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not defined ISCC (
    echo [WARNING] ไม่พบ Inno Setup 6 - ข้ามการสร้างตัวติดตั้ง
    echo ติดตั้งจาก https://jrsoftware.org/isdl.php แล้วรัน build_all.bat ใหม่
    echo EXE พร้อมใช้ที่ dist\SD8System\SD8System.exe
    goto :done
)
echo สร้างตัวติดตั้งด้วย Inno Setup...
"%ISCC%" installer.iss
if errorlevel 1 (
    echo [ERROR] Inno Setup compile ไม่สำเร็จ
    goto :fail
)
if not exist "Output\SD8System_Setup.exe" (
    echo [ERROR] ไม่พบ Output\SD8System_Setup.exe
    goto :fail
)

REM ---------- 10) MakeRelease: SHA-256 + update.json + ZIP ----------
echo.
echo สร้าง update.json และ ZIP...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0MakeRelease.ps1"
if errorlevel 1 (
    echo [ERROR] MakeRelease ไม่สำเร็จ
    goto :fail
)

:done
echo.
echo ============================================
echo   BUILD SUCCESS  -  Beta %APP_VER%
echo   EXE      : dist\SD8System\SD8System.exe
echo   Installer: Output\SD8System_Setup.exe
echo   Release  : release\SD8System_Beta_%APP_VER%.zip
echo ============================================
echo ถ้า Windows Defender/แอนตี้ไวรัสบล็อก .exe ให้ดูไฟล์ README_BUILD.txt หัวข้อ false positive
if /i not "%~1"=="nopause" pause
exit /b 0

:fail
echo.
echo ============================================
echo   BUILD FAILED
echo ============================================
if /i not "%~1"=="nopause" pause
exit /b 1

REM ---------- ลบโฟลเดอร์แบบ retry (OneDrive/AV ล็อกไฟล์) ----------
:safe_rmdir
set "_target=%~1"
if not exist "%_target%" goto :eof
set "_tries=0"
:safe_rmdir_loop
rmdir /s /q "%_target%" 2>nul
if not exist "%_target%" goto :eof
set /a _tries+=1
if %_tries% geq 5 (
    echo [ERROR] ลบ "%_target%" ไม่สำเร็จ - ปิด Explorer/แอนตี้ไวรัส หรือย้ายโปรเจกต์ออกจาก OneDrive
    goto :fail
)
timeout /t 2 /nobreak >nul
goto :safe_rmdir_loop
