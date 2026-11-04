# PowerShell Installer for Duy Bao - Word Tools
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Cài đặt Duy Bảo - Word Tools (Windows)"

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "       CÀI ĐẶT DUY BẢO - WORD TOOLS TRÊN WINDOWS" -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""

$targetDir = "$env:LOCALAPPDATA\DuyBaoWordTools"
$targetXml = "$targetDir\manifest.xml"
$addinId = "b785d0d8-5f2d-424a-9e12-427df18bb01a"

if (!(Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

Write-Host "[1/3] Đang chuẩn bị tệp cấu hình (manifest.xml)..." -ForegroundColor Yellow
$currentScriptManifest = Join-Path $PSScriptRoot "manifest.xml"
if (Test-Path $currentScriptManifest) {
    Copy-Item $currentScriptManifest -Destination $targetXml -Force
    Write-Host "      - Đã sao chép manifest.xml từ thư mục hiện tại." -ForegroundColor Gray
} else {
    Write-Host "      - Đang tải manifest.xml mới nhất từ GitHub..." -ForegroundColor Gray
    $url = "https://raw.githubusercontent.com/dbao8425/Thesis-Support-Tool-for-Word-macOS-_VN/main/manifest.xml"
    Invoke-WebRequest -Uri $url -OutFile $targetXml -UseBasicParsing
}

if (!(Test-Path $targetXml)) {
    Write-Host "[LỖI] Không tìm thấy hoặc không tải được manifest.xml!" -ForegroundColor Red
    Read-Host "Nhấn Enter để thoát..."
    exit 1
}

Write-Host "[2/3] Đang đăng ký Add-in vào Microsoft Word..." -ForegroundColor Yellow
$regPaths = @(
    "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer",
    "HKCU:\Software\Microsoft\Office\15.0\WEF\Developer"
)

foreach ($regPath in $regPaths) {
    if (!(Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    Set-ItemProperty -Path $regPath -Name $addinId -Value $targetXml -Type String -Force | Out-Null
}

Write-Host "[3/3] Đang làm mới bộ nhớ đệm Office..." -ForegroundColor Yellow
$cachePattern = "$env:LOCALAPPDATA\Microsoft\Office\16.0\WEF\$addinId*"
Get-Item $cachePattern -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "        🎉 CÀI ĐẶT THÀNH CÔNG CHO WINDOWS! 🎉" -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "HƯỚNG DẪN SỬ DỤNG:" -ForegroundColor Yellow
Write-Host "1. Hãy mở (hoặc khởi động lại) Microsoft Word trên máy tính."
Write-Host "2. Nhìn lên thanh Ribbon trên cùng, bạn sẽ thấy tab mới: 'Duy Bảo'."
Write-Host "3. Bấm vào nút 'Format Tables' để mở bảng công cụ."
Write-Host ""
Write-Host "* Lưu ý: Nếu Word đang mở sẵn trước đó, hãy tắt hẳn Word rồi mở lại." -ForegroundColor DarkGray
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""
Read-Host "Nhấn Enter để kết thúc..."
