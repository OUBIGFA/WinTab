param(
    [string]$WinTabExe = "E:\_Free code\WinTab\WinTab\bin\Release\net9.0-windows\win-x64\WinTab.exe",
    [int]$Iterations = 20,
    [int]$SampleMilliseconds = 1800,
    [int]$SampleIntervalMilliseconds = 30
)

$ErrorActionPreference = "Stop"

Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class ExplorerVisibilityProbe
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool GetLayeredWindowAttributes(IntPtr hwnd, out uint pcrKey, out byte pbAlpha, out uint pdwFlags);

    private const int GWL_EXSTYLE = -20;
    private const int WS_EX_LAYERED = 0x80000;

    public static IntPtr[] GetExplorerWindows()
    {
        var results = new List<IntPtr>();
        EnumWindows((hWnd, lParam) =>
        {
            var sb = new StringBuilder(256);
            GetClassName(hWnd, sb, sb.Capacity);
            if (string.Equals(sb.ToString(), "CabinetWClass", StringComparison.OrdinalIgnoreCase))
                results.Add(hWnd);
            return true;
        }, IntPtr.Zero);
        return results.ToArray();
    }

    public static bool IsLayeredAlphaZero(IntPtr hWnd)
    {
        var exStyle = GetWindowLong(hWnd, GWL_EXSTYLE);
        if ((exStyle & WS_EX_LAYERED) == 0)
            return false;

        uint colorKey;
        byte alpha;
        uint flags;
        if (!GetLayeredWindowAttributes(hWnd, out colorKey, out alpha, out flags))
            return false;

        return alpha == 0;
    }
}
"@

function Start-WinTab {
    Get-Process WinTab -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
    $process = Start-Process -FilePath $WinTabExe -PassThru
    Start-Sleep -Seconds 3
    if ($process.HasExited) {
        throw "WinTab exited during startup with code $($process.ExitCode)."
    }
    return $process
}

function Test-ExplorerInteractiveWindow {
    param(
        [int]$DurationMs,
        [int]$IntervalMs
    )

    $deadline = [DateTime]::UtcNow.AddMilliseconds($DurationMs)
    $maxInteractive = 0
    $snapshots = @()
    while ([DateTime]::UtcNow -lt $deadline) {
        $windows = [ExplorerVisibilityProbe]::GetExplorerWindows()
        $interactive = 0
        $layeredAlphaZero = 0
        foreach ($hWnd in $windows) {
            $visible = [ExplorerVisibilityProbe]::IsWindowVisible($hWnd)
            $iconic = [ExplorerVisibilityProbe]::IsIconic($hWnd)
            $alphaZero = [ExplorerVisibilityProbe]::IsLayeredAlphaZero($hWnd)
            if ($alphaZero) { $layeredAlphaZero++ }
            if ($visible -and -not $iconic -and -not $alphaZero) { $interactive++ }
        }

        if ($interactive -gt $maxInteractive) {
            $maxInteractive = $interactive
        }

        $snapshots += [pscustomobject]@{
            ExplorerWindows = $windows.Count
            InteractiveWindows = $interactive
            LayeredAlphaZero = $layeredAlphaZero
        }

        Start-Sleep -Milliseconds $IntervalMs
    }

    return [pscustomobject]@{
        MaxInteractiveWindows = $maxInteractive
        LastSnapshot = $snapshots[-1]
        Passed = ($maxInteractive -ge 1)
    }
}

$winTab = $null
$root = Join-Path ([System.IO.Path]::GetTempPath()) ("WinTabVisibilityProbe_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $root | Out-Null

try {
    $winTab = Start-WinTab

    for ($i = 1; $i -le $Iterations; $i++) {
        $folder = Join-Path $root ("case_{0:d3}" -f $i)
        New-Item -ItemType Directory -Path $folder | Out-Null
        Start-Process explorer.exe -ArgumentList "`"$folder`""
        Start-Sleep -Milliseconds 700

        # Also exercise taskbar/open-new-window code path.
        Start-Process explorer.exe
        Start-Sleep -Milliseconds 300

        $check = Test-ExplorerInteractiveWindow -DurationMs $SampleMilliseconds -IntervalMs $SampleIntervalMilliseconds
        $result = [pscustomobject]@{
            Iteration = $i
            MaxInteractiveWindows = $check.MaxInteractiveWindows
            LastInteractiveWindows = $check.LastSnapshot.InteractiveWindows
            LastLayeredAlphaZero = $check.LastSnapshot.LayeredAlphaZero
            Passed = $check.Passed
        }
        $result | ConvertTo-Json -Compress

        if (-not $check.Passed) {
            throw "No interactive Explorer window was available during sampling."
        }
    }

    "PASS: Explorer always had at least one interactive window during stress run."
}
finally {
    if ($winTab -and -not $winTab.HasExited) {
        Stop-Process -Id $winTab.Id -Force -ErrorAction SilentlyContinue
    }
}
