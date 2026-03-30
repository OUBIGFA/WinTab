# WinTab COM Registration Test Script
# Tests DelegateExecute COM registration for "Class not registered" issue
# 
# This script simulates how Windows 11 Start Menu and third-party apps
# check COM registration, which uses HKLM (machine-wide) context.

param(
    [switch]$FixMode,
    [string]$AppPath = "C:\Program Files\WinTab"
)

$ErrorActionPreference = "Stop"

# CLSID for WinTab DelegateExecute
$DelegateExecuteClsid = "{FD5BF2CD-0B24-4A80-9AF3-E40F9AFC0001}"
$ClsidPath = "Software\Classes\CLSID\$DelegateExecuteClsid"

# Colors for output
$Red = "Red"
$Green = "Green"
$Yellow = "Yellow"
$Cyan = "Cyan"
$White = "White"

function Write-TestHeader {
    param([string]$Title)
    Write-Host "`n$('=' * 60)" -ForegroundColor $Cyan
    Write-Host "  $Title" -ForegroundColor $Cyan
    Write-Host "$('=' * 60)" -ForegroundColor $Cyan
}

function Write-TestResult {
    param(
        [string]$Test,
        [bool]$Passed,
        [string]$Details = ""
    )
    $status = if ($Passed) { "PASS" } else { "FAIL" }
    $color = if ($Passed) { $Green } else { $Red }
    Write-Host "  [$status] $Test" -ForegroundColor $color
    if ($Details) {
        Write-Host "        $Details" -ForegroundColor $White
    }
    return $Passed
}

function Test-RegistryKey {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey,
        [Microsoft.Win32.RegistryView]$View = [Microsoft.Win32.RegistryView]::Default
    )
    
    try {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
        $key = $baseKey.OpenSubKey($SubKey, $false)
        if ($key) {
            $key.Close()
            return $true
        }
        return $false
    } catch {
        return $false
    }
}

function Test-RegistryValue {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey,
        [string]$ValueName,
        [string]$ExpectedValue = "",
        [Microsoft.Win32.RegistryView]$View = [Microsoft.Win32.RegistryView]::Default
    )
    
    try {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
        $key = $baseKey.OpenSubKey($SubKey, $false)
        if ($key) {
            $value = $key.GetValue($ValueName)
            $key.Close()
            
            if ([string]::IsNullOrEmpty($ExpectedValue)) {
                return $value -ne $null
            } else {
                return $value -eq $ExpectedValue
            }
        }
        return $false
    } catch {
        return $false
    }
}

function Get-RegistryValue {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey,
        [string]$ValueName,
        [Microsoft.Win32.RegistryView]$View = [Microsoft.Win32.RegistryView]::Default
    )
    
    try {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
        $key = $baseKey.OpenSubKey($SubKey, $false)
        if ($key) {
            $value = $key.GetValue($ValueName)
            $key.Close()
            return $value
        }
        return $null
    } catch {
        return $null
    }
}

# Track test results
$allPassed = $true
$testResults = @{
    HKLM_64_Exists = $false
    HKLM_64_Path = $false
    HKLM_64_Threading = $false
    HKLM_32_Exists = $false
    HKLM_32_Path = $false
    HKLM_32_Threading = $false
    HKCU_64_Exists = $false
    HKCU_64_Path = $false
    HKCU_64_Threading = $false
    HKCU_32_Exists = $false
    HKCU_32_Path = $false
    HKCU_32_Threading = $false
}

Write-TestHeader "WinTab DelegateExecute COM Registration Test"
Write-Host "  CLSID: $DelegateExecuteClsid" -ForegroundColor $White
Write-Host "  App Path: $AppPath" -ForegroundColor $White
Write-Host "  Test Mode: $(if ($FixMode) { 'FIX MODE' } else { 'DIAGNOSE MODE' })" -ForegroundColor $Yellow

# Test HKLM 64-bit (CRITICAL for Windows 11 Start Menu)
Write-TestHeader "HKLM (Machine-wide) - 64-bit Registry View"
Write-Host "  This is what Windows 11 Start Menu sees:" -ForegroundColor $Yellow

$result = Test-RegistryKey -Hive LocalMachine -SubKey $ClsidPath -View Registry64
$testResults.HKLM_64_Exists = $result
$allPassed = $allPassed -and $result
Write-TestResult "CLSID key exists" $result

if ($result) {
    $inprocPath = "$ClsidPath\InProcServer32"
    $expectedPath = "$AppPath\WinTab.ShellBridge.comhost.dll"
    
    $result = Test-RegistryKey -Hive LocalMachine -SubKey $inprocPath -View Registry64
    $testResults.HKLM_64_Path = $result
    $allPassed = $allPassed -and $result
    Write-TestResult "InProcServer32 key exists" $result
    
    if ($result) {
        $actualPath = Get-RegistryValue -Hive LocalMachine -SubKey $inprocPath -ValueName "" -View Registry64
        $result = $actualPath -eq $expectedPath
        $testResults.HKLM_64_Path = $result
        $allPassed = $allPassed -and $result
        Write-TestResult "comhost.dll path correct" $result "Expected: $expectedPath, Got: $actualPath"
        
        $threadingModel = Get-RegistryValue -Hive LocalMachine -SubKey $inprocPath -ValueName "ThreadingModel" -View Registry64
        $result = $threadingModel -eq "Apartment"
        $testResults.HKLM_64_Threading = $result
        $allPassed = $allPassed -and $result
        Write-TestResult "ThreadingModel = Apartment" $result "Got: $threadingModel"
    }
}

# Test HKLM 32-bit (for 32-bit apps)
Write-TestHeader "HKLM (Machine-wide) - 32-bit Registry View"
Write-Host "  This is what 32-bit third-party apps see:" -ForegroundColor $Yellow

$result = Test-RegistryKey -Hive LocalMachine -SubKey $ClsidPath -View Registry32
$testResults.HKLM_32_Exists = $result
$allPassed = $allPassed -and $result
Write-TestResult "CLSID key exists" $result

if ($result) {
    $inprocPath = "$ClsidPath\InProcServer32"
    $expectedPath = "$AppPath\x86\WinTab.ShellBridge.comhost.dll"
    
    $result = Test-RegistryKey -Hive LocalMachine -SubKey $inprocPath -View Registry32
    $testResults.HKLM_32_Path = $result
    $allPassed = $allPassed -and $result
    Write-TestResult "InProcServer32 key exists" $result
    
    if ($result) {
        $actualPath = Get-RegistryValue -Hive LocalMachine -SubKey $inprocPath -ValueName "" -View Registry32
        $result = $actualPath -eq $expectedPath
        $testResults.HKLM_32_Path = $result
        $allPassed = $allPassed -and $result
        Write-TestResult "comhost.dll path correct" $result "Expected: $expectedPath, Got: $actualPath"
        
        $threadingModel = Get-RegistryValue -Hive LocalMachine -SubKey $inprocPath -ValueName "ThreadingModel" -View Registry32
        $result = $threadingModel -eq "Apartment"
        $testResults.HKLM_32_Threading = $result
        $allPassed = $allPassed -and $result
        Write-TestResult "ThreadingModel = Apartment" $result "Got: $threadingModel"
    }
}

# Test HKCU 64-bit (user-only registration)
Write-TestHeader "HKCU (User-only) - 64-bit Registry View"
Write-Host "  This is what same-user Explorer sees:" -ForegroundColor $Yellow

$result = Test-RegistryKey -Hive CurrentUser -SubKey $ClsidPath -View Registry64
$testResults.HKCU_64_Exists = $result
Write-TestResult "CLSID key exists" $result

if ($result) {
    $inprocPath = "$ClsidPath\InProcServer32"
    $expectedPath = "$AppPath\WinTab.ShellBridge.comhost.dll"
    
    $result = Test-RegistryKey -Hive CurrentUser -SubKey $inprocPath -View Registry64
    $testResults.HKCU_64_Path = $result
    Write-TestResult "InProcServer32 key exists" $result
    
    if ($result) {
        $actualPath = Get-RegistryValue -Hive CurrentUser -SubKey $inprocPath -ValueName "" -View Registry64
        $result = $actualPath -eq $expectedPath
        $testResults.HKCU_64_Path = $result
        Write-TestResult "comhost.dll path correct" $result "Expected: $expectedPath, Got: $actualPath"
        
        $threadingModel = Get-RegistryValue -Hive CurrentUser -SubKey $inprocPath -ValueName "ThreadingModel" -View Registry64
        $result = $threadingModel -eq "Apartment"
        $testResults.HKCU_64_Threading = $result
        Write-TestResult "ThreadingModel = Apartment" $result "Got: $threadingModel"
    }
}

# Test HKCU 32-bit (user-only registration)
Write-TestHeader "HKCU (User-only) - 32-bit Registry View"
Write-Host "  This is what same-user 32-bit apps see:" -ForegroundColor $Yellow

$result = Test-RegistryKey -Hive CurrentUser -SubKey $ClsidPath -View Registry32
$testResults.HKCU_32_Exists = $result
Write-TestResult "CLSID key exists" $result

if ($result) {
    $inprocPath = "$ClsidPath\InProcServer32"
    $expectedPath = "$AppPath\x86\WinTab.ShellBridge.comhost.dll"
    
    $result = Test-RegistryKey -Hive CurrentUser -SubKey $inprocPath -View Registry32
    $testResults.HKCU_32_Path = $result
    Write-TestResult "InProcServer32 key exists" $result
    
    if ($result) {
        $actualPath = Get-RegistryValue -Hive CurrentUser -SubKey $inprocPath -ValueName "" -View Registry32
        $result = $actualPath -eq $expectedPath
        $testResults.HKCU_32_Path = $result
        Write-TestResult "comhost.dll path correct" $result "Expected: $expectedPath, Got: $actualPath"
        
        $threadingModel = Get-RegistryValue -Hive CurrentUser -SubKey $inprocPath -ValueName "ThreadingModel" -View Registry32
        $result = $threadingModel -eq "Apartment"
        $testResults.HKCU_32_Threading = $result
        Write-TestResult "ThreadingModel = Apartment" $result "Got: $threadingModel"
    }
}

# Summary
Write-TestHeader "Test Summary"

$criticalPassed = $testResults.HKLM_64_Exists -and $testResults.HKLM_64_Path -and $testResults.HKLM_64_Threading
$hkcuPassed = $testResults.HKCU_64_Exists -and $testResults.HKCU_64_Path -and $testResults.HKCU_64_Threading

Write-Host "  HKLM (Machine-wide) Registration:" -ForegroundColor $(if ($criticalPassed) { $Green } else { $Red })
Write-Host "    64-bit: $(if ($testResults.HKLM_64_Exists -and $testResults.HKLM_64_Path -and $testResults.HKLM_64_Threading) { 'COMPLETE' } else { 'MISSING/INCOMPLETE' })" -ForegroundColor $(if ($criticalPassed) { $Green } else { $Red })
Write-Host "    32-bit: $(if ($testResults.HKLM_32_Exists -and $testResults.HKLM_32_Path -and $testResults.HKLM_32_Threading) { 'COMPLETE' } else { 'MISSING/INCOMPLETE' })" -ForegroundColor $(if ($testResults.HKLM_32_Exists -and $testResults.HKLM_32_Path -and $testResults.HKLM_32_Threading) { $Green } else { $Yellow })

Write-Host "  HKCU (User-only) Registration:" -ForegroundColor $(if ($hkcuPassed) { $Green } else { $Yellow })
Write-Host "    64-bit: $(if ($testResults.HKCU_64_Exists -and $testResults.HKCU_64_Path -and $testResults.HKCU_64_Threading) { 'COMPLETE' } else { 'MISSING/INCOMPLETE' })" -ForegroundColor $(if ($hkcuPassed) { $Green } else { $Yellow })
Write-Host "    32-bit: $(if ($testResults.HKCU_32_Exists -and $testResults.HKCU_32_Path -and $testResults.HKCU_32_Threading) { 'COMPLETE' } else { 'MISSING/INCOMPLETE' })" -ForegroundColor $(if ($testResults.HKCU_32_Exists -and $testResults.HKCU_32_Path -and $testResults.HKCU_32_Threading) { $Green } else { $Yellow })

Write-Host "`n  Root Cause Analysis:" -ForegroundColor $Cyan
if (-not $criticalPassed) {
    Write-Host "    Windows 11 Start Menu and third-party apps CANNOT find the COM server" -ForegroundColor $Red
    Write-Host "    because it is only registered in HKCU (user scope), not HKLM (machine scope)." -ForegroundColor $Red
    Write-Host "    This causes 'Class not registered' error." -ForegroundColor $Red
} else {
    Write-Host "    COM server is properly registered in HKLM - should work correctly." -ForegroundColor $Green
}

if ($FixMode) {
    Write-Host "`n  FIX MODE: Registering COM server in HKLM..." -ForegroundColor $Green
    
    # Register in HKLM 64-bit
    Write-Host "  Registering HKLM 64-bit..." -ForegroundColor $Cyan
    reg add "HKLM\Software\Classes\CLSID\$DelegateExecuteClsid" /ve /d "WinTab Open Folder DelegateExecute" /f | Out-Null
    reg add "HKLM\Software\Classes\CLSID\$DelegateExecuteClsid\InProcServer32" /ve /d "$AppPath\WinTab.ShellBridge.comhost.dll" /f | Out-Null
    reg add "HKLM\Software\Classes\CLSID\$DelegateExecuteClsid\InProcServer32" /v "ThreadingModel" /d "Apartment" /f | Out-Null
    
    # Register in HKLM 32-bit
    Write-Host "  Registering HKLM 32-bit..." -ForegroundColor $Cyan
    reg add "HKLM\Software\Classes\CLSID\$DelegateExecuteClsid" /ve /d "WinTab Open Folder DelegateExecute" /f /reg:32 | Out-Null
    reg add "HKLM\Software\Classes\CLSID\$DelegateExecuteClsid\InProcServer32" /ve /d "$AppPath\x86\WinTab.ShellBridge.comhost.dll" /f /reg:32 | Out-Null
    reg add "HKLM\Software\Classes\CLSID\$DelegateExecuteClsid\InProcServer32" /v "ThreadingModel" /d "Apartment" /f /reg:32 | Out-Null
    
    Write-Host "  Registration complete. Restart Explorer for changes to take effect." -ForegroundColor $Green
    Write-Host "  Run: taskkill /f /im explorer.exe && start explorer.exe" -ForegroundColor $Yellow
}

Write-Host "`n"

# Return exit code
if ($allPassed) {
    exit 0
} else {
    exit 1
}
