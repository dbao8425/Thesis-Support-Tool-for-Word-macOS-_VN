@echo off
setlocal enabledelayedexpansion
title Cai dat Duy Bao - Word Tools cho Windows
color 0A

echo ========================================================
echo       CAI DAT DUY BAO - WORD TOOLS CHO WINDOWS
echo ========================================================
echo.

set "TARGET_DIR=%LOCALAPPDATA%\DuyBaoWordTools"
set "TARGET_XML=%TARGET_DIR%\manifest.xml"

if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

echo [1/3] Dang chuan bi tep manifest.xml...
if exist "%~dp0manifest.xml" (
    copy /y "%~dp0manifest.xml" "%TARGET_XML%" >nul
    echo       - Da sao chep manifest.xml tu thu muc hien tai.
) else (
    echo       - Dang tai manifest.xml tu GitHub...
    powershell -NoProfile -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile('https://raw.githubusercontent.com/dbao8425/Thesis-Support-Tool-for-Word-macOS-_VN/main/manifest.xml', '%TARGET_XML%')"
)

if not exist "%TARGET_XML%" (
    echo.
    echo [LOI] Khong the chuan bi manifest.xml!
    echo Vui long kiem tra ket noi Internet va thu lai.
    echo.
    pause
    exit /b 1
)

echo [2/3] Dang dang ky Add-in vao Microsoft Word...
reg add "HKCU\Software\Microsoft\Office.0\WEF\Developer" /v "b785d0d8-5f2d-424a-9e12-427df18bb01a" /t REG_SZ /d "%TARGET_XML%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office.0\WEF\Developer" /v "b785d0d8-5f2d-424a-9e12-427df18bb01a" /t REG_SZ /d "%TARGET_XML%" /f >nul 2>&1

echo [3/3] Dang lam moi bo nho dem Word...
powershell -NoProfile -Command "Remove-Item -Path ":LOCALAPPDATA\Microsoft\Office.0\WEF785d0d8-5f2d-424a-9e12-427df18bb01a*" -Recurse -Force -ErrorAction SilentlyContinue"

echo.
echo ========================================================
echo        CAI DAT THANH CONG CHO WINDOWS!
echo ========================================================
echo.
echo HUONG DAN SU DUNG:
echo 1. Hay mo (hoac khoi dong lai) Microsoft Word.
echo 2. Tren thanh Ribbon tren cung, ban se thay tab: "Duy Bao".
echo 3. Bam nut "Format Tables" de mo bang cong cu.
echo.
echo * Luu y: Neu Word dang mo, hay tat han Word roi mo lai.
echo ========================================================
echo.
pause
