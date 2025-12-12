# 一鍵打包腳本
# 執行此腳本自動完成打包流程

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "股票數據處理系統 - 自動打包工具" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# 檢查是否安裝 pyinstaller
Write-Host "[1/4] 檢查 PyInstaller..." -ForegroundColor Yellow
$pyinstaller = Get-Command pyinstaller -ErrorAction SilentlyContinue

if (-not $pyinstaller) {
    Write-Host "PyInstaller 未安裝，正在安裝..." -ForegroundColor Yellow
    pip install pyinstaller
    if ($LASTEXITCODE -ne 0) {
        Write-Host "安裝失敗！請手動執行：pip install pyinstaller" -ForegroundColor Red
        pause
        exit 1
    }
}
Write-Host "✓ PyInstaller 已就緒" -ForegroundColor Green
Write-Host ""

# 清理舊的打包檔案
Write-Host "[2/4] 清理舊檔案..." -ForegroundColor Yellow
if (Test-Path "dist") {
    Remove-Item -Recurse -Force "dist"
    Write-Host "✓ 已清理 dist 資料夾" -ForegroundColor Green
}
if (Test-Path "build") {
    Remove-Item -Recurse -Force "build"
    Write-Host "✓ 已清理 build 資料夾" -ForegroundColor Green
}
Write-Host ""

# 執行打包
Write-Host "[3/4] 開始打包..." -ForegroundColor Yellow
Write-Host "這可能需要幾分鐘，請耐心等候..." -ForegroundColor Yellow
pyinstaller stock_processor.spec

if ($LASTEXITCODE -ne 0) {
    Write-Host "✗ 打包失敗！" -ForegroundColor Red
    pause
    exit 1
}
Write-Host "✓ 打包完成" -ForegroundColor Green
Write-Host ""

# 整理發布檔案
Write-Host "[4/4] 整理發布檔案..." -ForegroundColor Yellow

# 創建發布資料夾
$releaseFolder = "release"
if (Test-Path $releaseFolder) {
    Remove-Item -Recurse -Force $releaseFolder
}
New-Item -ItemType Directory -Path $releaseFolder | Out-Null

# 複製必要檔案
Copy-Item "dist\股票數據處理器.exe" "$releaseFolder\" -ErrorAction SilentlyContinue
Copy-Item "執行股票處理.bat" "$releaseFolder\" -ErrorAction SilentlyContinue
Copy-Item "打包與部署指南.md" "$releaseFolder\使用說明.md" -ErrorAction SilentlyContinue
Copy-Item "Excel_VBA_代碼.vba" "$releaseFolder\" -ErrorAction SilentlyContinue

# 如果有 data 資料夾也複製
if (Test-Path "data") {
    Copy-Item -Recurse "data" "$releaseFolder\" -ErrorAction SilentlyContinue
}

Write-Host "✓ 發布檔案已整理至 release 資料夾" -ForegroundColor Green
Write-Host ""

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "打包完成！" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "發布檔案位置：$releaseFolder\" -ForegroundColor Cyan
Write-Host ""
Write-Host "包含檔案：" -ForegroundColor White
Write-Host "  - 股票數據處理器.exe" -ForegroundColor White
Write-Host "  - 執行股票處理.bat" -ForegroundColor White
Write-Host "  - Excel_VBA_代碼.vba" -ForegroundColor White
Write-Host "  - 使用說明.md" -ForegroundColor White
Write-Host ""
Write-Host "請將 release 資料夾中的所有檔案" -ForegroundColor Yellow
Write-Host "複製到沒有 Python 環境的電腦上使用" -ForegroundColor Yellow
Write-Host ""

# 開啟 release 資料夾
$open = Read-Host "要開啟 release 資料夾嗎？(Y/N)"
if ($open -eq "Y" -or $open -eq "y") {
    Invoke-Item $releaseFolder
}

Write-Host ""
Write-Host "按任意鍵結束..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
