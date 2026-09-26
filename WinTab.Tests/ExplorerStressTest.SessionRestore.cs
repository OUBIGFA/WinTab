using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.Managers;
using WinTab.Models;
using WinTab.WinAPI;

internal static partial class ExplorerStressTest
{
    private const string SessionStressFolder = "WinTabSessionStress";
    private const string RecycleBinPage = "shell:::{645FF040-5081-101B-9F08-00AA002F954E}";

    /// <summary>
    /// Restores saved groups into real Explorer windows this probe opens itself, through the watcher's own
    /// restore path. Windows that were open before the probe are hidden from the watcher, never receive a
    /// command, and must be unchanged at the end. The installed WinTab must be stopped first: it would merge
    /// the probe's windows into the user's. The watcher runs without merging, as with merging turned off.
    /// </summary>
    public static async Task<int> RunSessionRestoreAsync(string[] args)
    {
        if (Process.GetProcessesByName("WinTab").Length > 0)
        {
            Console.Error.WriteLine("Stop WinTab first: it would merge the probe's Explorer windows into existing windows.");
            return 2;
        }

        var stressRoot = Path.Combine(Path.GetTempPath(), SessionStressFolder);
        var root = Path.Combine(stressRoot, DateTime.Now.ToString("yyyyMMddHHmmssfff"));
        var a = CreateSessionFolder(root, "A");
        var b = CreateSessionFolder(root, "B");
        var c = CreateSessionFolder(root, "C");
        var start = CreateSessionFolder(root, "start");
        var debugLog = Path.Combine(root, "wintab-debug.log");
        Environment.SetEnvironmentVariable(ExplorerDebugLog.FileOverrideVariable, debugLog);

        // Explorer may show a hidden preloaded frame as the new window for this launch.
        var userFrames = ExplorerWindowDiscovery.GetAllExplorerWindows()
            .Where(ExplorerWindowDiscovery.IsShownExplorerWindow).ToHashSet();
        var userTabs = DescribeFrames(userFrames);
        Console.WriteLine($"User Explorer frames hidden from the probe: {userFrames.Count} ({userTabs.Values.Sum(tabs => tabs.Count)} tabs)");
        var failures = new List<string>();
        var ownedFrames = new HashSet<WindowIdentity>();
        ExplorerWatcher? watcher = null;
        try
        {
            watcher = new ExplorerWatcher();
            SetWatcherField(watcher, "_getExplorerWindows", (Func<IEnumerable<nint>>)(() =>
                ExplorerWindowDiscovery.GetAllExplorerWindows().Where(handle => !userFrames.Contains(handle)).ToArray()));
            var store = new ExplorerSessionStore(Path.Combine(root, "session.json"));
            SetWatcherField(watcher, "_sessionStore", store);
            watcher.StatusChanged += message => Console.WriteLine("  status: " + message);
            if (!await WaitForConditionAsync(() => watcher.IsShellReady, 10_000))
                throw new InvalidOperationException("The watcher did not connect to Explorer.");
            watcher.SetRestoreOnAnyFolder(false);
            watcher.SetRestoreTabs(true);
            // Filtering discovery alone does not filter ShellWindows callbacks. Explicitly exclude every
            // pre-existing user frame as well, so a delayed registration can never restore into one.
            var excludeUserFrame = typeof(ExplorerWatcher).GetMethod("ExcludeWindowFromSessionRestore",
                BindingFlags.Instance | BindingFlags.NonPublic)!;
            foreach (var userFrame in userFrames)
                excludeUserFrame.Invoke(watcher, [userFrame]);
            await Task.Delay(1_500);

            // 1. A normal launch replaces the start page with the whole saved group and selects the saved tab.
            await SaveSessionAsync(store, [a, b, c], 1);
            var frame = await OpenSessionProbeWindowAsync(userFrames, ownedFrames, null);
            await ExpectProbeTabsAsync(frame, ["A", "B", "C"], "B", failures, "normal launch");
            if (failures.Count > 0)
                throw new InvalidOperationException("Stopped after the first restoration failed.");
            await WaitForSessionRestoreIdleAsync(watcher);

            // 2. The restored window is the next saved group, with the tab the user selected last.
            if (IsSessionProbeFrame(frame, userFrames, root))
            {
                WinApi.TrySendMessage(frame, WinApi.WM_COMMAND, 0xA221, 3, 1_000);
                await ExpectProbeTabsAsync(frame, ["A", "B", "C"], "C", failures, "user selects C");
                await Task.Delay(1_500);
            }
            await CloseSessionProbeFrameAsync(frame, userFrames, root, ownedFrames, failures);
            await ExpectSavedSessionAsync(store, [a, b, c], 2, failures, "restored window closed");
            if (failures.Count > 0)
                throw new InvalidOperationException("Stopped after the restored window could not be saved.");
            if (!args.Contains("--quick", StringComparer.OrdinalIgnoreCase))
            {
                // A hidden preload may become the next window after this one closes.
                watcher.SetRestoreOnAnyFolder(true);
                await Task.Delay(6_000);
                await SaveSessionAsync(store, [a, b, c], 2);
                frame = await OpenSessionProbeWindowAsync(userFrames, ownedFrames, start);
                await ExpectProbeTabsAsync(frame, ["start", "A", "B", "C"], "start", failures, "any-folder explicit folder");
                if (failures.Count > 0) throw new InvalidOperationException("Stopped after any-folder restore failed.");
                await WaitForSessionRestoreIdleAsync(watcher);
                await CloseSessionProbeFrameAsync(frame, userFrames, root, ownedFrames, failures);
                await ExpectSavedSessionAsync(store, [start, a, b, c], 0, failures, "explicit-folder window closed");
                if (failures.Count > 0) throw new InvalidOperationException("Stopped after the folder window could not be saved.");

                // A plain launch, as from the taskbar icon, opens Explorer's start page. In any-folder mode that
                // page is what the user opened: it stays first and active, and the group follows it.
                await Task.Delay(6_000);
                await SaveSessionAsync(store, [a, b, c], 1);
                frame = await OpenSessionProbeWindowAsync(userFrames, ownedFrames, null);
                await ExpectStartPageKeptAsync(frame, ["A", "B", "C"], failures, "any-folder plain launch keeps the start page");
                if (failures.Count > 0) throw new InvalidOperationException("Stopped after any-folder plain launch failed.");
                await WaitForSessionRestoreIdleAsync(watcher);
                await CloseSessionProbeFrameAsync(frame, userFrames, root, ownedFrames, failures);
                // A frame can disappear before its asynchronous OnQuit/capture work has saved the group.
                // Wait for that save before injecting the next scenario's history, or the old close overwrites it.
                var keptGroupSaved = await WaitForConditionAsync(() => store.Snapshot is { } saved &&
                    saved.Locations.Length == 4 && saved.ActiveTabIndex == 0 &&
                    saved.Locations.Skip(1).Select(Helper.NormalizeLocation).SequenceEqual(new[] { a, b, c }, StringComparer.OrdinalIgnoreCase), 5_000);
                if (!keptGroupSaved)
                    throw new InvalidOperationException("The any-folder start-page window was not saved before the next scenario.");
                Console.WriteLine("PASS start-page window closed: initial page and restored group saved");

                watcher.SetRestoreOnAnyFolder(false);
                await SaveSessionAsync(store, [a, b], 0);
                frame = await OpenSessionProbeWindowAsync(userFrames, ownedFrames, start);
                await Task.Delay(4_000);
                await ExpectProbeTabsAsync(frame, ["start"], "start", failures, "normal-launch mode explicit folder", 1_000);
                await CloseSessionProbeFrameAsync(frame, userFrames, root, ownedFrames, failures);
                // Single-tab restore is off by default: a window with one tab does not replace the saved group.
                await ExpectSavedSessionAsync(store, [a, b], 0, failures, "single-tab window closed keeps the group");
                if (failures.Count > 0) throw new InvalidOperationException("Stopped after strict-mode restore failed.");

                var recycleBin = GetShellPageTitle(RecycleBinPage);
                await SaveSessionAsync(store, [a, RecycleBinPage, Path.Combine(root, "missing"), c], 1);
                frame = await OpenSessionProbeWindowAsync(userFrames, ownedFrames, null);
                await ExpectProbeTabsAsync(frame, ["A", recycleBin, "C"], recycleBin, failures, "built-in page and missing folder");
                await WaitForSessionRestoreIdleAsync(watcher);
                await CloseSessionProbeFrameAsync(frame, userFrames, root, ownedFrames, failures);
            }
        }
        catch (Exception exception)
        {
            failures.Add($"probe error: {exception.GetType().Name}: {exception.Message}");
        }
        finally
        {
            CloseSessionProbeFrames(userFrames, stressRoot, ownedFrames);
            watcher?.Dispose();
            ExplorerDebugLog.Complete(TimeSpan.FromSeconds(2));
            var now = DescribeFrames(userFrames);
            foreach (var (handle, tabs) in userTabs)
            {
                if (!now.TryGetValue(handle, out var current) || !current.SequenceEqual(tabs))
                    failures.Add($"user window {handle} changed: before=[{string.Join(" | ", tabs)}] after=[{string.Join(" | ", current ?? [])}]");
            }
            if (failures.Count == 0)
                Console.WriteLine($"PASS user windows unchanged: {userTabs.Count} frame(s)");
        }

        foreach (var failure in failures)
            Console.Error.WriteLine("FAIL " + failure);
        if (failures.Count > 0)
            DumpDebugLog(debugLog);
        else
            PrintRestoreTimeline(debugLog);
        if (!ownedFrames.Any(identity => identity.IsCurrent && ExplorerWindowDiscovery.IsFileExplorerWindow(identity.Handle)))
            TryDelete(root);
        else
            Console.Error.WriteLine($"Probe folders retained for an open test window: {root}");
        return failures.Count == 0 ? 0 : 1;
    }

    /// <summary>The restore steps with their timestamps, to see where a restore spends its time.</summary>
    private static void PrintRestoreTimeline(string debugLog)
    {
        if (!File.Exists(debugLog))
            return;
        using var stream = new FileStream(debugLog, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var reader = new StreamReader(stream);
        Console.WriteLine("Restore timeline:");
        while (reader.ReadLine() is { } line)
            if (line.Contains("Session restore") || line.Contains("OpenTab direct tab="))
                Console.WriteLine("  " + line);
    }

    private static async Task ExpectStartPageKeptAsync(nint frame, string[] restored, List<string> failures, string scenario)
    {
        var startPages = new[] { GetShellPageTitle("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"),
            GetShellPageTitle("shell:::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}") };
        var started = Environment.TickCount64;
        ExplorerTabAutomation.Tab[] tabs = [];
        bool Matches()
        {
            try { tabs = ExplorerTabAutomation.ReadSessionTabs(frame); }
            catch (Exception) { tabs = []; }
            return tabs.Length == restored.Length + 1 && startPages.Contains(tabs[0].Title, StringComparer.OrdinalIgnoreCase) &&
                tabs.Skip(1).Select(tab => tab.Title).SequenceEqual(restored, StringComparer.OrdinalIgnoreCase) &&
                tabs.Count(tab => tab.Selected) == 1 && tabs[0].Selected &&
                ExplorerWindowDiscovery.GetAllExplorerTabs(frame).Count() == tabs.Length;
        }
        if (await WaitForConditionAsync(Matches, 15_000))
        {
            // Stay a moment: a late close of the start page would show up here.
            await Task.Delay(1_500);
            if (Matches())
            {
                Console.WriteLine($"PASS {scenario}: [{string.Join(", ", tabs.Select(tab => tab.Title))}] active={tabs[0].Title} after {Environment.TickCount64 - started} ms");
                return;
            }
        }
        var actual = string.Join(", ", tabs.Select(tab => tab.Selected ? "*" + tab.Title : tab.Title));
        failures.Add($"{scenario}: expected [start page, {string.Join(", ", restored)}] with the start page active, got [{actual}]");
    }

    private static async Task WaitForSessionRestoreIdleAsync(ExplorerWatcher watcher)
    {
        var field = typeof(ExplorerWatcher).GetField("_sessionRestoresInProgress", BindingFlags.Instance | BindingFlags.NonPublic)!;
        if (!await WaitForConditionAsync(() => (int)field.GetValue(watcher)! == 0, 6_000))
            throw new InvalidOperationException("The restore transaction did not settle before the probe's next action.");
        await Task.Delay(300);
    }

    private static string CreateSessionFolder(string root, string name)
    {
        var path = Path.Combine(root, name);
        Directory.CreateDirectory(path);
        return path;
    }

    private static void SetWatcherField(ExplorerWatcher watcher, string name, object value) =>
        typeof(ExplorerWatcher).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(watcher, value);

    private static async Task<bool> WaitForConditionAsync(Func<bool> condition, int timeoutMs)
    {
        var deadline = Environment.TickCount64 + timeoutMs;
        while (Environment.TickCount64 < deadline)
        {
            if (condition())
                return true;
            await Task.Delay(50);
        }
        return condition();
    }

    private static async Task SaveSessionAsync(ExplorerSessionStore store, string[] locations, int active)
    {
        if (!await store.SaveAsync(new ExplorerSession { Locations = locations, ActiveTabIndex = active, OrderVerified = true }))
            throw new InvalidOperationException("The probe session could not be saved.");
    }

    /// <summary>Tab locations per frame, in a stable order, read through Explorer's own catalog.</summary>
    private static Dictionary<nint, List<string>> DescribeFrames(IEnumerable<nint> frames)
    {
        var wanted = frames.ToHashSet();
        return GetShellWindows().Where(window => wanted.Contains((nint)window.Hwnd))
            .GroupBy(window => (nint)window.Hwnd)
            .ToDictionary(group => group.Key, group => group
                .Select(window => string.IsNullOrEmpty(window.LocationUrl) ? window.Path : window.LocationUrl)
                .OrderBy(location => location, StringComparer.OrdinalIgnoreCase).ToList());
    }

    /// <summary>A frame the probe may command: not the user's, and every tab is a probe folder or a built-in page.</summary>
    private static bool IsSessionProbeFrame(nint frame, HashSet<nint> userFrames, string root)
    {
        if (frame == 0 || userFrames.Contains(frame) || !ExplorerWindowDiscovery.IsFileExplorerWindow(frame))
            return false;
        var tabs = GetShellWindows().Where(window => (nint)window.Hwnd == frame).ToArray();
        return tabs.Length > 0 && tabs.Length == ExplorerWindowDiscovery.GetAllExplorerTabs(frame).Count() &&
            tabs.Any(tab => Helper.NormalizeLocation(string.IsNullOrEmpty(tab.LocationUrl) ? tab.Path : tab.LocationUrl)
                .StartsWith(root + "\\", StringComparison.OrdinalIgnoreCase)) && tabs.All(tab =>
            {
                var location = Helper.NormalizeLocation(string.IsNullOrEmpty(tab.LocationUrl) ? tab.Path : tab.LocationUrl);
                return location.StartsWith(root + "\\", StringComparison.OrdinalIgnoreCase) ||
                    ExplorerSessionLocationPolicy.IsKnownShellPage(location);
            });
    }

    private static async Task<nint> OpenSessionProbeWindowAsync(HashSet<nint> userFrames,
        HashSet<WindowIdentity> ownedFrames, string? folder)
    {
        var shownBefore = ExplorerWindowDiscovery.GetAllExplorerWindows().Where(WinApi.IsWindowVisible).ToHashSet();
        if (folder == null)
            StartExplorerDefault();
        else
            StartExplorerNewWindow(folder);
        nint frame = 0;
        if (!await WaitForConditionAsync(() => (frame = ExplorerWindowDiscovery.GetAllExplorerWindows().FirstOrDefault(handle =>
                !userFrames.Contains(handle) && !shownBefore.Contains(handle) && WinApi.IsWindowVisible(handle))) != 0, 8_000))
            throw new InvalidOperationException("The probe's Explorer window did not open.");
        var identity = WindowIdentity.Capture(frame);
        if (!identity.IsCurrent)
            throw new InvalidOperationException("The probe could not identify its own Explorer window.");
        ownedFrames.Add(identity);
        // Restoration only proceeds in the foreground window, as when the user opens it.
        if (ExplorerNavigationAccess.ForegroundFrame() != frame)
            Helper.RestoreWindowToForeground(frame);
        Console.WriteLine($"  opened probe frame {frame} ({folder ?? "normal launch"})");
        return frame;
    }

    private static async Task ExpectProbeTabsAsync(nint frame, string[] titles, string selected, List<string> failures,
        string scenario, int timeoutMs = 15_000)
    {
        var started = Environment.TickCount64;
        ExplorerTabAutomation.Tab[] tabs = [];
        bool Matches()
        {
            try { tabs = ExplorerTabAutomation.ReadSessionTabs(frame); }
            catch (Exception) { tabs = []; }
            return tabs.Select(tab => tab.Title).SequenceEqual(titles, StringComparer.OrdinalIgnoreCase) &&
                tabs.Count(tab => tab.Selected) == 1 &&
                string.Equals(tabs.Single(tab => tab.Selected).Title, selected, StringComparison.OrdinalIgnoreCase) &&
                ExplorerWindowDiscovery.GetAllExplorerTabs(frame).Count() == titles.Length;
        }
        if (await WaitForConditionAsync(Matches, timeoutMs))
        {
            Console.WriteLine($"PASS {scenario}: [{string.Join(", ", titles)}] active={selected} after {Environment.TickCount64 - started} ms");
            return;
        }
        var actual = string.Join(", ", tabs.Select(tab => tab.Selected ? "*" + tab.Title : tab.Title));
        failures.Add($"{scenario}: expected [{string.Join(", ", titles)}] active={selected}, got [{actual}]");
    }

    private static async Task ExpectSavedSessionAsync(ExplorerSessionStore store, string[] locations, int active,
        List<string> failures, string scenario)
    {
        bool Matches() => store.Snapshot is { } session && session.ActiveTabIndex == active &&
            session.Locations.Select(Helper.NormalizeLocation).SequenceEqual(locations.Select(Helper.NormalizeLocation),
                StringComparer.OrdinalIgnoreCase);
        if (await WaitForConditionAsync(Matches, 5_000))
        {
            Console.WriteLine($"PASS {scenario}: saved {locations.Length} tab(s), active index {active}");
            return;
        }
        var snapshot = store.Snapshot;
        failures.Add($"{scenario}: expected saved [{string.Join(", ", locations.Select(Path.GetFileName))}] active={active}, " +
            $"got [{string.Join(", ", snapshot?.Locations.Select(Path.GetFileName) ?? [])}] active={snapshot?.ActiveTabIndex}");
    }

    private static async Task CloseSessionProbeFrameAsync(nint frame, HashSet<nint> userFrames, string root,
        HashSet<WindowIdentity> ownedFrames, List<string> failures)
    {
        if (!IsSessionProbeFrame(frame, userFrames, root) || !IsOwnedProbeFrame(frame, userFrames, root, ownedFrames))
        {
            failures.Add($"frame {frame} is not provably the probe's own window; it was left open");
            return;
        }
        // The frame's WM_CLOSE can close only its active tab on tabbed Explorer. SC_CLOSE is the
        // window-level close requested by the system menu (the same action as Alt+F4).
        WinApi.PostMessage(frame, 0x0112, 0xF060, 0);
        if (!await WaitForConditionAsync(() => !ExplorerWindowDiscovery.IsFileExplorerWindow(frame) || !WinApi.IsWindowVisible(frame), 3_000))
        {
            failures.Add($"probe frame {frame} did not close as a whole window");
            QuitOwnedFrame(frame, userFrames, root, ownedFrames);
        }
        if (!await WaitForConditionAsync(() => !ExplorerWindowDiscovery.IsFileExplorerWindow(frame) || !WinApi.IsWindowVisible(frame), 3_000))
            failures.Add($"probe frame {frame} did not close");
    }

    private static bool IsOwnedProbeFrame(nint frame, HashSet<nint> userFrames, string stressRoot,
        HashSet<WindowIdentity> ownedFrames)
    {
        if (userFrames.Contains(frame) || !ownedFrames.Any(identity => identity.Handle == frame && identity.IsCurrent) ||
            !ExplorerWindowDiscovery.IsFileExplorerWindow(frame))
            return false;
        var tabs = GetShellWindows().Where(window => (nint)window.Hwnd == frame).ToArray();
        return tabs.Length > 0 && tabs.Length == ExplorerWindowDiscovery.GetAllExplorerTabs(frame).Count() &&
            tabs.All(tab =>
            {
                var location = Helper.NormalizeLocation(string.IsNullOrEmpty(tab.LocationUrl) ? tab.Path : tab.LocationUrl);
                return string.IsNullOrEmpty(location) ||
                    location.StartsWith(stressRoot + "\\", StringComparison.OrdinalIgnoreCase) ||
                    ExplorerSessionLocationPolicy.IsKnownShellPage(location);
            });
    }

    private static void QuitOwnedFrame(nint frame, HashSet<nint> userFrames, string stressRoot,
        HashSet<WindowIdentity> ownedFrames)
    {
        if (!IsOwnedProbeFrame(frame, userFrames, stressRoot, ownedFrames)) return;
        var shellType = Type.GetTypeFromProgID("Shell.Application")!;
        var shell = Activator.CreateInstance(shellType)!;
        try
        {
            var windows = shellType.InvokeMember("Windows", BindingFlags.InvokeMethod, null, shell, []);
            if (windows == null) return;
            try
            {
                var tabs = new List<object>();
                foreach (var tab in (System.Collections.IEnumerable)windows)
                {
                    var hwnd = Convert.ToInt64(tab.GetType().InvokeMember("HWND", BindingFlags.GetProperty, null, tab, []));
                    if ((nint)hwnd == frame) tabs.Add(tab);
                }
                foreach (var tab in tabs)
                {
                    try { tab.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, tab, []); }
                    catch (System.Reflection.TargetInvocationException) { break; }
                }
            }
            finally { Marshal.ReleaseComObject(windows); }
        }
        finally { Marshal.ReleaseComObject(shell); }
    }

    /// <summary>Closes leftover probe windows only when their identities and locations match this run.</summary>
    private static void CloseSessionProbeFrames(HashSet<nint> userFrames, string stressRoot,
        HashSet<WindowIdentity> ownedFrames)
    {
        foreach (var identity in ownedFrames.ToArray())
            if (identity.IsCurrent && IsOwnedProbeFrame(identity.Handle, userFrames, stressRoot, ownedFrames))
                QuitOwnedFrame(identity.Handle, userFrames, stressRoot, ownedFrames);
    }

    private static string GetShellPageTitle(string page)
    {
        var shellType = Type.GetTypeFromProgID("Shell.Application")!;
        var shell = Activator.CreateInstance(shellType)!;
        try
        {
            var folder = shellType.InvokeMember("NameSpace", BindingFlags.InvokeMethod, null, shell, [page]);
            return Convert.ToString(folder?.GetType().InvokeMember("Title", BindingFlags.GetProperty, null, folder, [])) ?? page;
        }
        finally
        {
            ReleaseComObject(shell);
        }
    }
}
