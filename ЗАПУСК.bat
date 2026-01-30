@echo off
chcp 65001 >nul
title Obrabotka dannyx 112 i Google Sheets
color 0A

echo ============================================================
echo    ZAPUSK PRILOZHENIYA
echo ============================================================
echo.

cd /d "%~dp0"

echo Aktivaciya virtualnogo okruzheniya...
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
    echo OK
) else (
    echo Virtualnoe okruzhenie ne naydeno
)

echo.
echo Zapusk programmy...
echo.

python process_data_app.py

if errorlevel 1 (
    echo.
    echo OSHIBKA!
    pause
)

exit
