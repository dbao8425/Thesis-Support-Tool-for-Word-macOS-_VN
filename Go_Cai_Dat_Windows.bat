@echo off
setlocal
title Go cai dat Duy Bao - Word Tools
color 0C

echo ========================================================
echo      GO CAI DAT DUY BAO - WORD TOOLS TREN WINDOWS
echo ========================================================
echo.

echo [1/2] Dang xoa dang ky Add-in khoi Microsoft Word...
reg delete "HKCU\Software\Microsoft\Office.0\WEF\Developer" /v "b785d0d8-5f2d-424a-9e12-427df18bb01a" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office.0\WEF\Developer" /v "b785d0d8-5f2d-424a-9e12-427df18bb01a" /f >nul 2>&1

echo [2/2] Dang don dep cac tep cai dat...
if exist "%LOCALAPPDATA%\DuyBaoWordTools" rd /s /q "%LOCALAPPDATA%\DuyBaoWordTools" >nul 2>&1
powershell -NoProfile -Command "Remove-Item -Path ":LOCALAPPDATA\Microsoft\Office.0\WEF785d0d8-5f2d-424a-9e12-427df18bb01a*" -Recurse -Force -ErrorAction SilentlyContinue"

echo.
echo ========================================================
echo      DA GO BO THANH CONG KHOI MICROSOFT WORD!
echo ========================================================
echo Khi ban mo lai Microsoft Word, tab "Duy Bao" se khong con nua.
echo.
pause
