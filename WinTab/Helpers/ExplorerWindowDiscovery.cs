using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using WinTab.WinAPI;

namespace WinTab.Helpers;

public static class ExplorerWindowDiscovery
{
    public static bool IsFileExplorerTab(nint tab)
    {
        return tab != 0 && WinApi.IsWindowHasClassName(tab, "ShellTabWindowClass");
    }

    public static bool IsFileExplorerWindow(nint window)
    {
        return window != 0 && WinApi.IsWindowHasClassName(window, "CabinetWClass");
    }

    public static bool IsFileExplorerForeground(out nint foregroundWindow)
    {
        foregroundWindow = WinApi.GetForegroundWindow();
        return IsFileExplorerWindow(foregroundWindow);
    }

    public static nint GetAnotherExplorerWindow(nint currentWindow)
    {
        return currentWindow == 0
            ? WinApi.FindWindow("CabinetWClass", null)
            : GetAllExplorerWindows().FirstOrDefault(window => window != currentWindow);
    }

    public static Task<nint> ListenForNewExplorerWindowAsync(IReadOnlyCollection<nint> currentWindows, int searchTimeMs = 1000)
    {
        var knownWindows = CreateKnownHandleSet(currentWindows);
        return Helper.DoUntilNotDefaultAsync(() =>
                GetAllExplorerWindows()
                    .FirstOrDefault(window => IsUnknownHandle(window, knownWindows)),
            searchTimeMs);
    }

    public static nint ListenForNewExplorerTab(IReadOnlyCollection<nint> currentTabs, int searchTimeMs = 1000)
    {
        var knownTabs = CreateKnownHandleSet(currentTabs);
        return Helper.DoUntilNotDefault(() =>
                GetAllExplorerTabs()
                    .FirstOrDefault(tab => IsUnknownHandle(tab, knownTabs)),
            searchTimeMs);
    }

    public static Task<nint> ListenForNewExplorerTabAsync(IReadOnlyCollection<nint> currentTabs, int searchTimeMs = 1000)
    {
        var knownTabs = CreateKnownHandleSet(currentTabs);
        return Helper.DoUntilNotDefaultAsync(() =>
                GetAllExplorerTabs()
                    .FirstOrDefault(tab => IsUnknownHandle(tab, knownTabs)),
            searchTimeMs);
    }

    public static Task<nint> ListenForNewExplorerTabAsync(nint window, IReadOnlyCollection<nint> currentTabs, int searchTimeMs = 1000)
    {
        var knownTabs = CreateKnownHandleSet(currentTabs);
        return Helper.DoUntilNotDefaultAsync(() =>
                GetAllExplorerTabs(window)
                    .FirstOrDefault(tab => IsUnknownHandle(tab, knownTabs)),
            searchTimeMs);
    }

    public static List<nint> GetAllExplorerTabs()
    {
        var tabs = new List<nint>();

        foreach (var window in GetAllExplorerWindows())
            tabs.AddRange(GetAllExplorerTabs(window));

        return tabs;
    }

    public static IEnumerable<nint> GetAllExplorerTabs(nint window)
    {
        return WinApi.FindAllWindowsEx("ShellTabWindowClass", window);
    }

    public static IEnumerable<nint> GetAllExplorerWindows()
    {
        return WinApi.FindAllWindowsEx("CabinetWClass");
    }

    public static Process? GetMainExplorerProcess()
    {
        Process? best = null;
        var windowsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        var expectedPath = System.IO.Path.Combine(windowsFolder, "explorer.exe");
        var bestStart = DateTime.MaxValue;

        foreach (var hWnd in WinApi.FindAllWindowsEx("Shell_TrayWnd"))
        {
            if (WinApi.GetWindowThreadProcessId(hWnd, out var pid) <= 0)
                continue;

            var processPath = WinApi.GetProcessPath((int)pid);
            if (!string.Equals(processPath, expectedPath, StringComparison.OrdinalIgnoreCase))
                continue;

            try
            {
                var proc = Process.GetProcessById((int)pid);
                if (proc.StartTime < bestStart)
                {
                    bestStart = proc.StartTime;
                    best = proc;
                }
            }
            catch
            {
                // The process can terminate between the window scan and Process lookup.
            }
        }

        return best;
    }

    private static HashSet<nint>? CreateKnownHandleSet(IReadOnlyCollection<nint> handles)
    {
        return handles.Count == 0 ? null : new HashSet<nint>(handles);
    }

    private static bool IsUnknownHandle(nint handle, HashSet<nint>? knownHandles)
    {
        return knownHandles == null || !knownHandles.Contains(handle);
    }
}
