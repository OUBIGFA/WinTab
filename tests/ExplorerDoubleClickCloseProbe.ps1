param(
    [string]$WinTabExe = "E:\_Free code\WinTab\WinTab\bin\Release\net9.0-windows\win-x64\WinTab.exe",
    [int]$TimeoutSeconds = 20
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class ExplorerDoubleClickProbe
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    public delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

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

    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    public const uint MOUSEEVENTF_LEFTUP = 0x0004;
    public const uint WM_COMMAND = 0x0111;
    public const uint KEYEVENTF_KEYUP = 0x0002;
    public const uint SWP_SHOWWINDOW = 0x0040;
    public const byte VK_CONTROL = 0x11;
    public const byte VK_T = 0x54;
    public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);

    public static IntPtr[] GetExplorerWindows()
    {
        var result = new List<IntPtr>();
        EnumWindows((hWnd, lParam) =>
        {
            var sb = new StringBuilder(256);
            GetClassName(hWnd, sb, sb.Capacity);
            if (string.Equals(sb.ToString(), "CabinetWClass", StringComparison.OrdinalIgnoreCase) && IsWindowVisible(hWnd))
                result.Add(hWnd);
            return true;
        }, IntPtr.Zero);
        return result.ToArray();
    }

    public static int CountExplorerTabs(IntPtr window)
    {
        var count = 0;
        var handle = IntPtr.Zero;
        do
        {
            handle = FindWindowEx(window, handle, "ShellTabWindowClass", null);
            if (handle != IntPtr.Zero)
                count++;
        } while (handle != IntPtr.Zero);

        return count;
    }

    public static void DoubleClick(int x, int y)
    {
        DoubleClickWithDrift(x, y, 0, 0, 80);
    }

    public static void DoubleClickWithDrift(int x, int y, int dx, int dy, int delayMs)
    {
        SetCursorPos(x, y);
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(delayMs);
        SetCursorPos(x + dx, y + dy);
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
    }

    public static void SendCtrlT()
    {
        keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
        keybd_event(VK_T, 0, 0, UIntPtr.Zero);
        keybd_event(VK_T, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }

    public static void OpenNewTab(IntPtr window)
    {
        var activeTab = FindWindowEx(window, IntPtr.Zero, "ShellTabWindowClass", null);
        if (activeTab != IntPtr.Zero)
            SendMessage(activeTab, WM_COMMAND, new IntPtr(0xA21B), IntPtr.Zero);
    }

    public static void BringExplorerToTestPosition(IntPtr window)
    {
        SetWindowPos(window, HWND_TOPMOST, 80, 80, 1000, 720, SWP_SHOWWINDOW);
        SetForegroundWindow(window);
        System.Threading.Thread.Sleep(100);
        SetWindowPos(window, HWND_NOTOPMOST, 80, 80, 1000, 720, SWP_SHOWWINDOW);
        SetForegroundWindow(window);
    }
}
"@

function Start-WinTab {
    Get-Process WinTab -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 600
    $process = Start-Process -FilePath $WinTabExe -PassThru
    Start-Sleep -Seconds 3
    if ($process.HasExited) {
        throw "WinTab exited during startup with code $($process.ExitCode)."
    }
    $process.Refresh()
    if ($process.MainWindowHandle -ne [IntPtr]::Zero) {
        [ExplorerDoubleClickProbe]::ShowWindow($process.MainWindowHandle, 0) | Out-Null
    }
    $process
}

function Wait-ExplorerWindow {
    param([long[]]$Exclude = @())

    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        $windows = [ExplorerDoubleClickProbe]::GetExplorerWindows() | Where-Object { $Exclude -notcontains $_.ToInt64() }
        if ($windows.Count -gt 0) {
            return $windows[0]
        }
        Start-Sleep -Milliseconds 100
    }

    return [IntPtr]::Zero
}

function Get-FirstTabTitlePoint {
    param([IntPtr]$Window)

    try {
        $root = [System.Windows.Automation.AutomationElement]::FromHandle($Window)
        $condition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::TabItem)
        $items = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
        foreach ($item in $items) {
            $rect = $item.Current.BoundingRectangle
            if (-not $rect.IsEmpty -and $rect.Width -ge 24 -and $rect.Height -ge 12) {
                return [pscustomobject]@{
                    X = [int]($rect.Left + [Math]::Min($rect.Width / 2, 120))
                    Y = [int]($rect.Top + ($rect.Height / 2))
                    Source = "UIA"
                }
            }
        }
    }
    catch {
    }

    [ExplorerDoubleClickProbe+RECT]$rect = New-Object ExplorerDoubleClickProbe+RECT
    [ExplorerDoubleClickProbe]::GetWindowRect($Window, [ref]$rect) | Out-Null
    return [pscustomobject]@{
        X = $rect.Left + 180
        Y = $rect.Top + 18
        Source = "fallback"
    }
}

function Get-BlankTitleBarPoint {
    param([IntPtr]$Window)

    try {
        $root = [System.Windows.Automation.AutomationElement]::FromHandle($Window)
        $condition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::TabItem)
        $items = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
        foreach ($item in $items) {
            $rect = $item.Current.BoundingRectangle
            if (-not $rect.IsEmpty -and $rect.Width -ge 24 -and $rect.Height -ge 12) {
                return [pscustomobject]@{
                    X = [int]($rect.Right + 220)
                    Y = [int]($rect.Top + ($rect.Height / 2))
                    Source = "UIA"
                }
            }
        }
    }
    catch {
    }

    [ExplorerDoubleClickProbe+RECT]$rect = New-Object ExplorerDoubleClickProbe+RECT
    [ExplorerDoubleClickProbe]::GetWindowRect($Window, [ref]$rect) | Out-Null
    return [pscustomobject]@{
        X = $rect.Left + 520
        Y = $rect.Top + 18
        Source = "fallback"
    }
}

function Wait-TabCount {
    param(
        [IntPtr]$Window,
        [int]$Minimum
    )

    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        $count = [ExplorerDoubleClickProbe]::CountExplorerTabs($Window)
        if ($count -ge $Minimum) {
            return $count
        }

        Start-Sleep -Milliseconds 100
    }

    return [ExplorerDoubleClickProbe]::CountExplorerTabs($Window)
}

$winTab = $null
$root = Join-Path ([System.IO.Path]::GetTempPath()) ("WinTabDoubleClickProbe_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $root | Out-Null

try {
    Get-Process WinTab -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 600

    $before = @([ExplorerDoubleClickProbe]::GetExplorerWindows() | ForEach-Object { $_.ToInt64() })
    $folder = Join-Path $root "single_tab"
    New-Item -ItemType Directory -Path $folder | Out-Null
    Start-Process explorer.exe -ArgumentList "`"$folder`""

    $window = Wait-ExplorerWindow -Exclude $before
    if ($window -eq [IntPtr]::Zero) {
        throw "No Explorer window was created for the single-tab test."
    }

    $winTab = Start-WinTab
    Start-Sleep -Seconds 1
    [ExplorerDoubleClickProbe]::BringExplorerToTestPosition($window)
    [ExplorerDoubleClickProbe]::SetForegroundWindow($window) | Out-Null
    Start-Sleep -Milliseconds 300
    [ExplorerDoubleClickProbe]::OpenNewTab($window)
    $tabsBeforeMulti = Wait-TabCount -Window $window -Minimum 2
    if ($tabsBeforeMulti -lt 2) {
        throw "Expected at least two tabs before multi-tab double-click, got $tabsBeforeMulti."
    }

    [ExplorerDoubleClickProbe]::SetForegroundWindow($window) | Out-Null
    Start-Sleep -Milliseconds 300
    [ExplorerDoubleClickProbe]::BringExplorerToTestPosition($window)
    $multiPoint = Get-FirstTabTitlePoint $window
    [ExplorerDoubleClickProbe]::DoubleClickWithDrift($multiPoint.X, $multiPoint.Y, 3, 1, 30)
    Start-Sleep -Seconds 2

    $stillVisibleAfterMulti = [ExplorerDoubleClickProbe]::GetExplorerWindows() | Where-Object { $_.ToInt64() -eq $window.ToInt64() }
    $tabsAfterMulti = if ($stillVisibleAfterMulti.Count -gt 0) { [ExplorerDoubleClickProbe]::CountExplorerTabs($window) } else { 0 }
    $multiResult = [pscustomobject]@{
        Window = $window.ToInt64()
        TabsBefore = $tabsBeforeMulti
        TabsAfter = $tabsAfterMulti
        ClickPoint = "$($multiPoint.X),$($multiPoint.Y)"
        ClickPointSource = $multiPoint.Source
        WindowStillVisible = ($stillVisibleAfterMulti.Count -gt 0)
        Passed = ($stillVisibleAfterMulti.Count -gt 0 -and $tabsAfterMulti -eq ($tabsBeforeMulti - 1))
    }
    $multiResult | ConvertTo-Json -Compress

    if (-not $multiResult.Passed) {
        throw "Double-click with slight pointer drift did not close exactly one tab in a multi-tab Explorer window."
    }

    Start-Sleep -Seconds 2
    $visibleAfterTwoTabClose = [ExplorerDoubleClickProbe]::IsWindowVisible($window)
    $minimizedAfterTwoTabClose = [ExplorerDoubleClickProbe]::IsIconic($window)
    $twoTabRecoveryResult = [pscustomobject]@{
        Window = $window.ToInt64()
        WindowVisible = $visibleAfterTwoTabClose
        WindowMinimized = $minimizedAfterTwoTabClose
        Tabs = [ExplorerDoubleClickProbe]::CountExplorerTabs($window)
        Passed = ($visibleAfterTwoTabClose -and -not $minimizedAfterTwoTabClose)
    }
    $twoTabRecoveryResult | ConvertTo-Json -Compress

    if (-not $twoTabRecoveryResult.Passed) {
        throw "Explorer was hidden or minimized after closing one of two tabs."
    }

    $afterTwoTabFolder = Join-Path $root "after_two_tab_close"
    New-Item -ItemType Directory -Path $afterTwoTabFolder | Out-Null
    Start-Process explorer.exe -ArgumentList "`"$afterTwoTabFolder`""
    Start-Sleep -Seconds 3
    $afterTwoTabOpenResult = [pscustomobject]@{
        Window = $window.ToInt64()
        WindowVisible = [ExplorerDoubleClickProbe]::IsWindowVisible($window)
        WindowMinimized = [ExplorerDoubleClickProbe]::IsIconic($window)
        Tabs = [ExplorerDoubleClickProbe]::CountExplorerTabs($window)
        Passed = ([ExplorerDoubleClickProbe]::IsWindowVisible($window) -and -not [ExplorerDoubleClickProbe]::IsIconic($window))
    }
    $afterTwoTabOpenResult | ConvertTo-Json -Compress

    if (-not $afterTwoTabOpenResult.Passed) {
        throw "Opening a folder after closing one of two tabs left Explorer hidden or minimized."
    }

    [ExplorerDoubleClickProbe]::SetForegroundWindow($window) | Out-Null
    Start-Sleep -Milliseconds 300
    [ExplorerDoubleClickProbe]::BringExplorerToTestPosition($window)
    $blankPoint = Get-BlankTitleBarPoint $window
    $tabsBeforeBlank = [ExplorerDoubleClickProbe]::CountExplorerTabs($window)
    [ExplorerDoubleClickProbe]::DoubleClickWithDrift($blankPoint.X, $blankPoint.Y, 2, 1, 30)
    Start-Sleep -Seconds 2
    $stillVisibleAfterBlank = [ExplorerDoubleClickProbe]::GetExplorerWindows() | Where-Object { $_.ToInt64() -eq $window.ToInt64() }
    $tabsAfterBlank = if ($stillVisibleAfterBlank.Count -gt 0) { [ExplorerDoubleClickProbe]::CountExplorerTabs($window) } else { 0 }
    $blankResult = [pscustomobject]@{
        Window = $window.ToInt64()
        TabsBefore = $tabsBeforeBlank
        TabsAfter = $tabsAfterBlank
        ClickPoint = "$($blankPoint.X),$($blankPoint.Y)"
        ClickPointSource = $blankPoint.Source
        WindowStillVisible = ($stillVisibleAfterBlank.Count -gt 0)
        Passed = ($stillVisibleAfterBlank.Count -gt 0 -and $tabsAfterBlank -eq $tabsBeforeBlank)
    }
    $blankResult | ConvertTo-Json -Compress

    if (-not $blankResult.Passed) {
        throw "Double-click on blank Explorer title bar area was incorrectly handled as a tab close."
    }

    $beforeReopen = @([ExplorerDoubleClickProbe]::GetExplorerWindows() | ForEach-Object { $_.ToInt64() })
    $reopenFolder = Join-Path $root "after_final_close"
    New-Item -ItemType Directory -Path $reopenFolder | Out-Null
    Start-Process explorer.exe -ArgumentList "`"$reopenFolder`""
    $reopenedWindow = Wait-ExplorerWindow -Exclude $beforeReopen
    $reopenResult = [pscustomobject]@{
        Window = if ($reopenedWindow -ne [IntPtr]::Zero) { $reopenedWindow.ToInt64() } else { 0 }
        WindowStillVisible = ($reopenedWindow -ne [IntPtr]::Zero)
        Passed = ($reopenedWindow -ne [IntPtr]::Zero)
    }
    $reopenResult | ConvertTo-Json -Compress

    if (-not $reopenResult.Passed) {
        throw "Opening a folder after closing the final Explorer tab did not create a visible Explorer window."
    }

    "PASS: Drifted tab double-click closes one of two tabs, the remaining Explorer stays visible, blank title bar is ignored, and later folder opens remain visible."
}
finally {
    if ($winTab -and -not $winTab.HasExited) {
        Stop-Process -Id $winTab.Id -Force -ErrorAction SilentlyContinue
    }
}
