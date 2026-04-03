param(
    [string]$LogPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Log {
    param([string]$Message)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[${timestamp}] $Message"
    Write-Output $line
    if ($script:ResolvedLogPath) {
        Add-Content -LiteralPath $script:ResolvedLogPath -Value $line -Encoding UTF8
    }
}

function Ensure-Admin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Administrator privileges are required."
    }
}

function Test-RegKeyExists {
    param(
        [string]$KeyPath,
        [string]$View = ""
    )

    $suffix = if ($View) { " /reg:$View" } else { "" }
    $command = "reg query `"$KeyPath`"$suffix >nul 2>nul"
    cmd.exe /c $command | Out-Null
    return $LASTEXITCODE -eq 0
}

function Export-RegKeyIfExists {
    param(
        [string]$KeyPath,
        [string]$DestinationPath,
        [string]$View = ""
    )

    if (-not (Test-RegKeyExists -KeyPath $KeyPath -View $View)) {
        Write-Log "[skip missing] $KeyPath"
        return
    }

    $args = @("export", $KeyPath, $DestinationPath, "/y")
    if ($View) {
        $args += "/reg:$View"
    }

    Write-Log ("reg " + ($args -join " "))
    & reg.exe @args 2>&1 | ForEach-Object { Write-Log $_ }
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to export $KeyPath"
    }
}

function Remove-RegKeyIfExists {
    param(
        [string]$KeyPath,
        [string]$View = ""
    )

    if (-not (Test-RegKeyExists -KeyPath $KeyPath -View $View)) {
        Write-Log "[skip missing] $KeyPath"
        return
    }

    $args = @("delete", $KeyPath, "/f")
    if ($View) {
        $args += "/reg:$View"
    }

    Write-Log ("reg " + ($args -join " "))
    & reg.exe @args 2>&1 | ForEach-Object { Write-Log $_ }
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to delete $KeyPath"
    }
}

function Remove-RegValueIfExists {
    param(
        [string]$KeyPath,
        [string]$ValueName
    )

    $queryArgs = @("query", $KeyPath)
    if ([string]::IsNullOrEmpty($ValueName)) {
        $queryArgs += "/ve"
    }
    else {
        $queryArgs += @("/v", $ValueName)
    }

    & reg.exe @queryArgs *> $null
    if ($LASTEXITCODE -ne 0) {
        $label = if ([string]::IsNullOrEmpty($ValueName)) { "(Default)" } else { $ValueName }
        Write-Log "[skip missing value] $KeyPath [$label]"
        return
    }

    $args = @("delete", $KeyPath)
    if ([string]::IsNullOrEmpty($ValueName)) {
        $args += "/ve"
    }
    else {
        $args += @("/v", $ValueName)
    }
    $args += "/f"
    Write-Log ("reg " + ($args -join " "))
    & reg.exe @args 2>&1 | ForEach-Object { Write-Log $_ }
    if ($LASTEXITCODE -ne 0) {
        $label = if ([string]::IsNullOrEmpty($ValueName)) { "(Default)" } else { $ValueName }
        throw "Failed to delete $KeyPath [$label]"
    }
}

function Remove-EmptyUserClassesKey {
    param([string]$SubKeyPath)

    $fullPath = "Registry::HKEY_CURRENT_USER\Software\Classes\$SubKeyPath"
    if (-not (Test-Path $fullPath)) {
        return
    }

    $item = Get-Item $fullPath -ErrorAction SilentlyContinue
    if ($null -eq $item) {
        return
    }

    $valueNames = $item.GetValueNames()
    $subKeys = $item.GetSubKeyNames()
    if ($valueNames.Count -eq 0 -and $subKeys.Count -eq 0) {
        Remove-Item -LiteralPath $fullPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Removed empty key $fullPath"
    }
}

function Remove-WinTabMuiCacheEntries {
    $muiCachePath = "Registry::HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache"
    if (-not (Test-Path $muiCachePath)) {
        return
    }

    $props = (Get-ItemProperty $muiCachePath).PSObject.Properties |
        Where-Object {
            $_.Name -notmatch '^PS' -and
            (
                $_.Name -match 'WinTab' -or
                ($_.Value -is [string] -and $_.Value -match 'WinTab')
            )
        }

    foreach ($prop in $props) {
        Remove-ItemProperty -LiteralPath $muiCachePath -Name $prop.Name -ErrorAction SilentlyContinue
        Write-Log "Removed MuiCache value $($prop.Name)"
    }
}

function Remove-WinTabClsidAllViews {
    param([string]$Clsid)

    foreach ($view in [Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32) {
        foreach ($hive in [Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryHive]::LocalMachine) {
            try {
                $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
                $classesKey = $baseKey.OpenSubKey("Software\\Classes\\CLSID", $true)
                if ($null -ne $classesKey) {
                    $classesKey.DeleteSubKeyTree($Clsid, $false)
                    $classesKey.Dispose()
                    Write-Log "Deleted CLSID $Clsid from $hive $view"
                }
                $baseKey.Dispose()
            }
            catch {
                Write-Log ("Failed deleting CLSID {0} from {1} {2}: {3}" -f $Clsid, $hive, $view, $_.Exception.Message)
            }
        }
    }
}

function Repair-FileExplorerTaskbarShortcut {
    $shortcutPath = Join-Path $env:APPDATA "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\File Explorer.lnk"
    $targetPath = "$env:WINDIR\explorer.exe"

    if (-not (Test-Path $shortcutPath)) {
        Write-Log "[skip missing] $shortcutPath"
        return
    }

    try {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        if ([string]::IsNullOrWhiteSpace($shortcut.TargetPath) -or -not (Test-Path $shortcut.TargetPath)) {
            $shortcut.TargetPath = $targetPath
            $shortcut.Arguments = ""
            $shortcut.WorkingDirectory = "$env:WINDIR"
            $shortcut.IconLocation = "$env:WINDIR\explorer.exe,0"
            $shortcut.Save()
            Write-Log "Repaired taskbar File Explorer shortcut"
        }
        else {
            Write-Log "Taskbar File Explorer shortcut already valid"
        }
    }
    catch {
        Write-Log "Failed to inspect or repair taskbar shortcut: $($_.Exception.Message)"
    }
}

Ensure-Admin

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ResolvedLogPath = if ($LogPath) {
    $LogPath
} else {
    Join-Path $scriptDir ("explorer-system-repair-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".log")
}

$backupDir = Join-Path $scriptDir ("explorer-system-repair-backup-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

Write-Log "Backup directory: $backupDir"

$backupTargets = @(
    @{ Key = "HKCU\Software\Classes\Folder"; File = "HKCU-Classes-Folder.reg"; View = "" },
    @{ Key = "HKCU\Software\Classes\Directory"; File = "HKCU-Classes-Directory.reg"; View = "" },
    @{ Key = "HKCU\Software\Classes\Drive"; File = "HKCU-Classes-Drive.reg"; View = "" },
    @{ Key = "HKCU\Software\WinTab"; File = "HKCU-Software-WinTab.reg"; View = "" },
    @{ Key = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.lnk"; File = "HKCU-Explorer-FileExts-lnk.reg"; View = "" },
    @{ Key = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run"; File = "HKCU-Run.reg"; View = "" },
    @{ Key = "HKLM\SOFTWARE\Classes\CLSID\{FD5BF2CD-0B24-4A80-9AF3-E40F9AFC0001}"; File = "HKLM-CLSID-WinTab-64.reg"; View = "64" },
    @{ Key = "HKLM\SOFTWARE\Classes\CLSID\{FD5BF2CD-0B24-4A80-9AF3-E40F9AFC0001}"; File = "HKLM-CLSID-WinTab-32.reg"; View = "32" }
)

foreach ($target in $backupTargets) {
    Export-RegKeyIfExists -KeyPath $target.Key -DestinationPath (Join-Path $backupDir $target.File) -View $target.View
}

$verbs = @("open", "explore", "opennewwindow")
$classes = @("Folder", "Directory", "Drive")

foreach ($cls in $classes) {
    foreach ($verb in $verbs) {
        Remove-RegKeyIfExists -KeyPath "HKCU\Software\Classes\$cls\shell\$verb"
    }

    Remove-RegValueIfExists -KeyPath "HKCU\Software\Classes\$cls\shell" -ValueName ""
    Remove-EmptyUserClassesKey -SubKeyPath "$cls\shell"
    Remove-EmptyUserClassesKey -SubKeyPath $cls
}

Remove-WinTabClsidAllViews -Clsid "{FD5BF2CD-0B24-4A80-9AF3-E40F9AFC0001}"
Remove-RegValueIfExists -KeyPath "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" -ValueName "WinTab"
Remove-RegKeyIfExists -KeyPath "HKCU\Software\WinTab"
Remove-RegKeyIfExists -KeyPath "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.lnk"

Remove-WinTabMuiCacheEntries
Repair-FileExplorerTaskbarShortcut

Write-Log "Resetting lightweight Explorer user state"
New-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Force | Out-Null
Set-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name LaunchTo -Type DWord -Value 2
Set-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name SeparateProcess -Type DWord -Value 0
Remove-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name HubMode -ErrorAction SilentlyContinue

$recentDirs = @(
    (Join-Path $env:APPDATA "Microsoft\Windows\Recent\AutomaticDestinations"),
    (Join-Path $env:APPDATA "Microsoft\Windows\Recent\CustomDestinations")
)
foreach ($dir in $recentDirs) {
    if (Test-Path $dir) {
        Get-ChildItem -LiteralPath $dir -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        Write-Log "Cleared $dir"
    }
}

Write-Log "Restarting Explorer"
Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Start-Process explorer.exe

Write-Log "WinTab Explorer cleanup completed"
Write-Log "If Explorer UI is still broken after reboot, the remaining issue is outside active WinTab hooks and should be handled with a fresh user profile test or Windows in-place repair."
