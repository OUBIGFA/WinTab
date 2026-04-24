param(
    [string]$WinTabExe = "E:\_Free code\WinTab\WinTab\bin\Release\net9.0-windows\win-x64\WinTab.exe",
    [int]$Iterations = 12,
    [int]$SampleMilliseconds = 2500,
    [int]$SampleIntervalMilliseconds = 10,
    [switch]$ValidateLocation
)

$ErrorActionPreference = "Stop"

Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class ExplorerWindowProbe
{
    public const int GWL_EXSTYLE = -20;
    public const int WS_EX_LAYERED = 0x80000;
    public const int LWA_ALPHA = 0x2;
    public const int DWMWA_CLOAKED = 14;

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    public static extern bool GetLayeredWindowAttributes(IntPtr hwnd, out uint crKey, out byte bAlpha, out uint dwFlags);

    [DllImport("dwmapi.dll")]
    public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out int pvAttribute, int cbAttribute);

    public static IntPtr[] GetExplorerWindows()
    {
        var result = new List<IntPtr>();
        EnumWindows((hWnd, lParam) =>
        {
            var sb = new StringBuilder(256);
            GetClassName(hWnd, sb, sb.Capacity);
            if (string.Equals(sb.ToString(), "CabinetWClass", StringComparison.OrdinalIgnoreCase))
                result.Add(hWnd);
            return true;
        }, IntPtr.Zero);
        return result.ToArray();
    }

    public static bool IsUserVisibleExplorerWindow(IntPtr hWnd)
    {
        if (!IsWindowVisible(hWnd))
            return false;

        int cloaked;
        if (DwmGetWindowAttribute(hWnd, DWMWA_CLOAKED, out cloaked, sizeof(int)) == 0 && cloaked != 0)
            return false;

        var exStyle = GetWindowLong(hWnd, GWL_EXSTYLE);
        if ((exStyle & WS_EX_LAYERED) != 0)
        {
            uint colorKey;
            byte alpha;
            uint flags;
            if (GetLayeredWindowAttributes(hWnd, out colorKey, out alpha, out flags) && (flags & LWA_ALPHA) != 0 && alpha == 0)
                return false;
        }

        RECT rect;
        if (!GetWindowRect(hWnd, out rect))
            return false;

        var width = rect.Right - rect.Left;
        var height = rect.Bottom - rect.Top;
        if (width <= 20 || height <= 20)
            return false;

        if (rect.Right < -1000 || rect.Bottom < -1000 || rect.Left > 20000 || rect.Top > 20000)
            return false;

        return true;
    }
}
"@

function Get-ExplorerWindowSet {
    $set = @{}
    foreach ($handle in [ExplorerWindowProbe]::GetExplorerWindows()) {
        $set[$handle.ToInt64()] = $true
    }
    $set
}

function Start-WinTab {
    Get-Process WinTab -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 600
    $process = Start-Process -FilePath $WinTabExe -PassThru
    Start-Sleep -Seconds 3
    if ($process.HasExited) {
        throw "WinTab exited during startup with code $($process.ExitCode)."
    }
    $process
}

function Convert-ExplorerLocationUrl {
    param([string]$LocationUrl)

    if ([string]::IsNullOrWhiteSpace($LocationUrl)) {
        return ""
    }

    if ($LocationUrl.StartsWith("file:///", [System.StringComparison]::OrdinalIgnoreCase)) {
        $path = [System.Uri]::UnescapeDataString($LocationUrl.Substring(8)).Replace("/", "\")
        return $path
    }

    return [System.Uri]::UnescapeDataString($LocationUrl)
}

function Test-ExplorerLocation {
    param([string]$ExpectedPath)

    $expected = [System.IO.Path]::GetFullPath($ExpectedPath).TrimEnd("\")
    $shell = New-Object -ComObject Shell.Application
    $windows = $null
    try {
        $windows = $shell.Windows()
        $count = $windows.Count
        for ($i = 0; $i -lt $count; $i++) {
            $window = $null
            try {
                $window = $windows.Item($i)
                if (-not ($window.FullName -like "*\explorer.exe")) {
                    continue
                }

                $actual = Convert-ExplorerLocationUrl $window.LocationURL
                if ([string]::IsNullOrWhiteSpace($actual)) {
                    continue
                }

                $actual = [System.IO.Path]::GetFullPath($actual).TrimEnd("\")
                if ([string]::Equals($actual, $expected, [System.StringComparison]::OrdinalIgnoreCase)) {
                    return $true
                }
            }
            catch {
            }
        }
    }
    finally {
        if ($windows -ne $null) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($windows) | Out-Null
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
    }

    return $false
}

$winTab = Start-WinTab
$root = Join-Path ([System.IO.Path]::GetTempPath()) ("WinTabFlashProbe_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $root | Out-Null

try {
    $baselineFolder = Join-Path $root "baseline"
    New-Item -ItemType Directory -Path $baselineFolder | Out-Null
    Start-Process explorer.exe -ArgumentList "`"$baselineFolder`""
    Start-Sleep -Seconds 3

    $failures = @()
    for ($i = 1; $i -le $Iterations; $i++) {
        $target = Join-Path $root ("case_" + $i.ToString("000"))
        New-Item -ItemType Directory -Path $target | Out-Null

        $before = Get-ExplorerWindowSet
        $visibleSamples = 0
        $visibleHandles = New-Object System.Collections.Generic.HashSet[long]

        Start-Process explorer.exe -ArgumentList "`"$target`""
        $deadline = [Environment]::TickCount + $SampleMilliseconds
        while ([Environment]::TickCount -lt $deadline) {
            foreach ($handle in [ExplorerWindowProbe]::GetExplorerWindows()) {
                $id = $handle.ToInt64()
                if ($before.ContainsKey($id)) {
                    continue
                }
                if ([ExplorerWindowProbe]::IsUserVisibleExplorerWindow($handle)) {
                    $visibleSamples++
                    [void]$visibleHandles.Add($id)
                }
            }
            Start-Sleep -Milliseconds $SampleIntervalMilliseconds
        }

        $targetMatched = if ($ValidateLocation) { Test-ExplorerLocation $target } else { $null }

        $result = [pscustomobject]@{
            Iteration = $i
            Target = $target
            VisibleSamples = $visibleSamples
            VisibleHandles = @($visibleHandles)
            TargetMatched = $targetMatched
            Passed = ($visibleSamples -eq 0 -and (-not $ValidateLocation -or $targetMatched))
        }

        $result | ConvertTo-Json -Compress
        if (-not $result.Passed) {
            $failures += $result
        }

        Start-Sleep -Milliseconds 200
    }

    if ($failures.Count -gt 0) {
        throw "Explorer flash detected in $($failures.Count) of $Iterations iterations."
    }

    "PASS: No user-visible transient Explorer window detected in $Iterations iterations."
}
finally {
    if ($winTab -and -not $winTab.HasExited) {
        Stop-Process -Id $winTab.Id -Force -ErrorAction SilentlyContinue
    }
}
