@echo off
chcp 65001 > nul
title อัปเดตฐานข้อมูลหนังสือ + ส่งขึ้นเว็บอัตโนมัติ

echo.
echo ============================================================
echo     ระบบอัปเดตฐานข้อมูลหนังสือเรียน (Auto Update)
echo     Scrape - บันทึก - ส่งขึ้น GitHub - เสร็จ!
echo ============================================================
echo.

REM --- ตรวจสอบ Python ---
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] ไม่พบ Python! กรุณาติดตั้ง Python ก่อน
    echo         ดาวน์โหลดที่: https://www.python.org/downloads/
    echo         (ติ๊ก "Add Python to PATH" ด้วยนะครับ)
    pause
    exit /b 1
)

REM --- ตรวจสอบ Git ---
git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] ไม่พบ Git! กรุณาติดตั้ง Git ก่อน
    echo         ดาวน์โหลดที่: https://git-scm.com/download/win
    pause
    exit /b 1
)

REM --- ติดตั้ง Library ที่จำเป็น ---
echo [1/3] กำลังตรวจสอบ Library...
pip install -q -r requirements.txt
echo       Library พร้อมใช้งาน!
echo.

REM --- รัน Scraper ---
echo [2/3] กำลังดึงข้อมูลหนังสือจากเว็บ... (อาจใช้เวลา 5-10 นาที)
echo       กรุณาอย่าปิดหน้าต่างนี้
echo.
python update_data.py
if errorlevel 1 (
    echo.
    echo [ERROR] การดึงข้อมูลล้มเหลว กรุณาตรวจสอบ:
    echo         - เชื่อมต่อเครือข่ายโรงเรียนอยู่หรือไม่?
    echo         - เซิร์ฟเวอร์ 202.29.173.190 ใช้งานได้ไหม?
    pause
    exit /b 1
)
echo.
echo       ดึงข้อมูลเสร็จแล้ว!
echo.

REM --- Git Push ---
echo [3/3] กำลังส่งข้อมูลขึ้น GitHub...
for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set TODAY=%%c-%%b-%%a
git add textbooks.xlsx
git commit -m "auto: อัปเดตฐานข้อมูลหนังสือ %date% %time:~0,5%"
git push

if errorlevel 1 (
    echo.
    echo [WARNING] Git push อาจมีปัญหา กรุณาตรวจสอบ:
    echo           - มี Internet (ไม่ใช่เฉพาะ Intranet) หรือไม่?
    echo           - Git ถูก Login GitHub แล้วหรือยัง?
) else (
    echo.
    echo ============================================================
    echo  สำเร็จ! ข้อมูลหนังสือชุดใหม่ถูกส่งขึ้นเว็บแล้ว
    echo  เว็บไซต์สำหรับคุณครูจะอัปเดตเองภายใน 1-2 นาที
    echo ============================================================
)

echo.
pause
