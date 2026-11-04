@echo off
chcp 65001 >nul
title Go cai dat Duy Bao - Word Tools
color 0C

echo ========================================================
echo      GỠ CÀI ĐẶT DUY BẢO - WORD TOOLS TRÊN WINDOWS
echo ========================================================
echo.

echo [1/2] Đang xóa đăng ký Add-in khỏi Microsoft Word...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "b785d0d8-5f2d-424a-9e12-427df18bb01a" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\15.0\WEF\Developer" /v "b785d0d8-5f2d-424a-9e12-427df18bb01a" /f >nul 2>&1

echo [2/2] Đang dọn dẹp các tệp cài đặt...
rd /s /q "%LOCALAPPDATA%\DuyBaoWordTools" >nul 2>&1
del /q /s "%LOCALAPPDATA%\Microsoft\Office\16.0\WEF\b785d0d8-5f2d-424a-9e12-427df18bb01a*" >nul 2>&1

echo.
echo ========================================================
echo      ĐÃ GỠ BỎ THÀNH CÔNG KHỎI MICROSOFT WORD!
echo ========================================================
echo Khi bạn mở lại Microsoft Word, tab "Duy Bảo" sẽ biến mất hoàn toàn.
echo.
pause
