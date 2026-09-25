using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WinTab.WinAPI;

namespace WinTab.Helpers;

public static class ExplorerWindowDiscovery
{
    public static bool IsFileExplorerWindow(nint window)
    {
        return window != 0 && WinApi.IsWindowHasClassName(window, "CabinetWClass");
    }

    /// <summary>
    /// An Explorer window the shell has actually shown. Windows 11 preloads hidden frames for later
    /// folder opens; those are not user windows and must never receive merged tabs.
    /// </summary>
    public static bool IsShownExplorerWindow(nint window)
    {
        return IsFileExplorerWindow(window) && WinApi.IsWindowVisible(window);
    }

    public static bool IsFileExplorerForeground(out nint foregroundWindow)
    {
        foregroundWindow = WinApi.GetForegroundWindow();
        return IsFileExplorerWindow(foregroundWindow);
    }

    public static Task<nint> ListenForNewExplorerWindowAsync(IReadOnlyCollection<nint> currentWindows, int searchTimeMs = 1000, CancellationToken cancellationToken = default)
    {
        var knownWindows = CreateKnownHandleSet(currentWindows);
        return Helper.DoUntilNotDefaultAsync(() =>
                GetAllExplorerWindows()
                    .FirstOrDefault(window => IsUnknownHandle(window, knownWindows)),
            searchTimeMs, cancellationToken: cancellationToken);
    }

    public static Task<nint> ListenForNewExplorerTabAsync(nint window, IReadOnlyCollection<nint> currentTabs, int searchTimeMs = 1000, CancellationToken cancellationToken = default)
    {
        var knownTabs = new HashSet<nint>(currentTabs);
        return Helper.DoUntilNotDefaultAsync(() => GetUniqueNewExplorerTab(window, knownTabs),
            searchTimeMs, cancellationToken: cancellationToken);
    }

    internal static nint GetUniqueNewExplorerTab(nint window, HashSet<nint> knownTabs)
    {
        var tabs = GetAllExplorerTabs(window).ToArray();
        if (!knownTabs.IsSubsetOf(tabs) || tabs.Length > knownTabs.Count + 1)
            throw new InvalidOperationException("Explorer tabs changed concurrently; no new tab can be safely claimed.");
        return tabs.FirstOrDefault(tab => !knownTabs.Contains(tab));
    }

    public static IEnumerable<nint> GetAllExplorerTabs(nint window)
    {
        return WinApi.FindAllWindowsEx("ShellTabWindowClass", window);
    }

    /// <summary>
    /// The frame's tabs as one consistent snapshot in z-order, the active tab first, or null while the frame
    /// keeps changing. Explorer moves tab windows in the z-order when it adds or activates a tab, and a sibling
    /// walk that spans such a move lists a tab twice or misses one.
    /// </summary>
    public static nint[]? GetStableExplorerTabs(nint window, int maxCount) =>
        ReadStable(() => GetAllExplorerTabs(window).Take(maxCount).ToArray());

    /// <summary>
    /// Two identical consecutive walks without a repeated entry are taken as a snapshot: a walk torn by a
    /// z-order move differs from the walk after it.
    /// </summary>
    internal static T[]? ReadStable<T>(Func<T[]> read, int attempts = 4) where T : notnull
    {
        var previous = read();
        for (var attempt = 1; attempt < attempts; attempt++)
        {
            var current = read();
            if (current.SequenceEqual(previous) && current.Distinct().Count() == current.Length)
                return current;
            previous = current;
        }
        return null;
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
