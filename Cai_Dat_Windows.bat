@echo off
chcp 65001 >nul
title Cai dat Duy Bao - Word Tools cho Windows
color 0A

echo ========================================================
echo       CÀI ĐẶT DUY BẢO - WORD TOOLS TRÊN WINDOWS
echo ========================================================
echo.

set "TARGET_DIR=%LOCALAPPDATA%\DuyBaoWordTools"
set "TARGET_XML=%TARGET_DIR%\manifest.xml"

if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

echo [1/3] Đang chuẩn bị tệp cấu hình (manifest.xml)...
if exist "%~dp0manifest.xml" (
    copy /y "%~dp0manifest.xml" "%TARGET_XML%" >nul
    echo       - Đã sao chép manifest.xml từ thư mục hiện tại.
) else (
    echo       - Đang tải manifest.xml mới nhất từ GitHub...
    curl -s -L "https://raw.githubusercontent.com/dbao8425/Thesis-Support-Tool-for-Word-macOS-_VN/main/manifest.xml" -o "%TARGET_XML%"
)

if not exist "%TARGET_XML%" (
    echo [LỖI] Không tìm thấy hoặc không tải được manifest.xml!
    echo Vui lòng kiểm tra lại kết nối Internet và thử lại.
    echo.
    pause
    exit /b 1
)

echo [2/3] Đang đăng ký Add-in vào Microsoft Word...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "b785d0d8-5f2d-424a-9e12-427df18bb01a" /t REG_SZ /d "%TARGET_XML%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\15.0\WEF\Developer" /v "b785d0d8-5f2d-424a-9e12-427df18bb01a" /t REG_SZ /d "%TARGET_XML%" /f >nul 2>&1

echo [3/3] Đang làm mới bộ nhớ đệm Office...
del /q /s "%LOCALAPPDATA%\Microsoft\Office\16.0\WEF\b785d0d8-5f2d-424a-9e12-427df18bb01a*" >nul 2>&1

echo.
echo ========================================================
echo        🎉 CÀI ĐẶT THÀNH CÔNG CHO WINDOWS! 🎉
echo ========================================================
echo.
echo HƯỚNG DẪN SỬ DỤNG:
echo 1. Hãy mở (hoặc khởi động lại) Microsoft Word trên máy tính.
echo 2. Nhìn lên thanh Ribbon trên cùng, bạn sẽ thấy tab mới: "Duy Bảo".
echo 3. Bấm vào nút "Format Tables" để mở bảng công cụ.
echo.
echo * Lưu ý: Nếu Word đang mở sẵn trước đó, hãy tắt hẳn Word rồi mở lại.
echo ========================================================
echo.
pause
