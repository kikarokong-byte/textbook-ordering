@echo off
chcp 65001 > nul
title Textbook Auto-Update

echo.
echo ============================================================
echo   [Step 1/3] Checking Python and Git...
echo ============================================================

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found!
    echo Please install Python from: https://www.python.org/downloads/
    echo Don't forget to check "Add Python to PATH"
    pause
    exit /b 1
)

git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git not found!
    echo Please install Git from: https://git-scm.com/download/win
    pause
    exit /b 1
)

echo [OK] Python and Git are ready.
echo.

echo ============================================================
echo   [Step 2/3] Installing required libraries...
echo ============================================================
pip install -q -r requirements.txt
echo [OK] Libraries installed.
echo.

echo ============================================================
echo   [Step 3a/3] Scraping textbook data... (5-10 minutes)
echo   Please wait. Do NOT close this window.
echo ============================================================
echo.
python update_data.py
if errorlevel 1 (
    echo.
    echo [ERROR] Scraping failed. Please check:
    echo   - Are you connected to the school network?
    echo   - Is server 202.29.173.190 accessible?
    pause
    exit /b 1
)
echo.
echo [OK] Scraping complete!
echo.

echo ============================================================
echo   [Step 3b/3] Pushing to GitHub...
echo ============================================================
git add textbooks.xlsx
git commit -m "auto: update textbooks database"
git push

if errorlevel 1 (
    echo.
    echo [WARNING] Git push may have failed. Check above for details.
    echo Make sure you have internet access and GitHub is authenticated.
) else (
    echo.
    echo ============================================================
    echo   SUCCESS! Database updated and pushed to GitHub.
    echo   The teacher website will update automatically in ~2 min.
    echo ============================================================
)

echo.
pause
