# WinTab COM Registration Fix Script
# Applies HKLM COM registration to fix "Class not registered" error
# Run as Administrator

param(
    [string]$AppPath = "C:\Program Files\WinTab"
)

$ErrorActionPreference = "Stop"

$DelegateExecuteClsid = "{FD5BF2CD-0B24-4A80-9AF3-E40F9AFC0001}"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  WinTab COM Registration Fix" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "App Path: $AppPath`n" -ForegroundColor White

# Check if running as admin
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "ERROR: This script must be run as Administrator" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'`n" -ForegroundColor Yellow
    exit 1
}

# Verify app path exists
if (-not (Test-Path $AppPath)) {
    Write-Host "WARNING: App path does not exist: $AppPath" -ForegroundColor Yellow
    Write-Host "Using default path anyway...`n" -ForegroundColor Yellow
}

Write-Host "Registering COM server in HKLM (machine-wide)...`n" -ForegroundColor Green

# HKLM 64-bit
Write-Host "  [1/6] HKLM 64-bit CLSID key..." -ForegroundColor Cyan
reg add "HKLM\Software\Classes\CLSID\$DelegateExecuteClsid" /ve /d "WinTab Open Folder DelegateExecute" /f | Out-Null

Write-Host "  [2/6] HKLM 64-bit InProcServer32..." -ForegroundColor Cyan
reg add "HKLM\Software\Classes\CLSID\$DelegateExecuteClsid\InProcServer32" /ve /d "$AppPath\WinTab.ShellBridge.comhost.dll" /f | Out-Null

Write-Host "  [3/6] HKLM 64-bit ThreadingModel..." -ForegroundColor Cyan
reg add "HKLM\Software\Classes\CLSID\$DelegateExecuteClsid\InProcServer32" /v "ThreadingModel" /d "Apartment" /f | Out-Null

# HKLM 32-bit
Write-Host "  [4/6] HKLM 32-bit CLSID key..." -ForegroundColor Cyan
reg add "HKLM\Software\Classes\CLSID\$DelegateExecuteClsid" /ve /d "WinTab Open Folder DelegateExecute" /f /reg:32 | Out-Null

Write-Host "  [5/6] HKLM 32-bit InProcServer32..." -ForegroundColor Cyan
reg add "HKLM\Software\Classes\CLSID\$DelegateExecuteClsid\InProcServer32" /ve /d "$AppPath\x86\WinTab.ShellBridge.comhost.dll" /f /reg:32 | Out-Null

Write-Host "  [6/6] HKLM 32-bit ThreadingModel..." -ForegroundColor Cyan
reg add "HKLM\Software\Classes\CLSID\$DelegateExecuteClsid\InProcServer32" /v "ThreadingModel" /d "Apartment" /f /reg:32 | Out-Null

Write-Host "`nRegistration complete!`n" -ForegroundColor Green

Write-Host "To apply changes, restart Explorer:" -ForegroundColor Yellow
Write-Host "  taskkill /f /im explorer.exe && start explorer.exe`n" -ForegroundColor White

Write-Host "Or restart your computer.`n" -ForegroundColor Yellow

exit 0
