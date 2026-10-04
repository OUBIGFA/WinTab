using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.WinAPI;

/// <summary>Opt-in desktop probe: the caller supplies a disposable single-tab Explorer window.</summary>
internal static class ExplorerDoubleClickNativeTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("explorer double-click returns through real activation history", Run);
    }

    private static async Task Run()
    {
        if (!long.TryParse(Environment.GetEnvironmentVariable("WINTAB_TEST_EXPLORER_HWND"), out var handle))
            throw new TestSkippedException("Requires a disposable Explorer window in WINTAB_TEST_EXPLORER_HWND");
        var window = (nint)handle;
        Check.That(ExplorerWindowDiscovery.IsFileExplorerWindow(window), "The supplied window must be Explorer.");
        var identity = WindowIdentity.Capture(window);
        Check.Equal(1, ExplorerTabAutomation.ReadSessionTabs(window).Length, "The owned window must start with one tab.");
        Helper.RestoreWindowToForeground(window);
        for (var count = 2; count <= 4; count++)
        {
            Check.That(identity.IsCurrent && WinApi.GetForegroundWindow() == window, "The owned window must hold foreground.");
            Check.That(WinApi.PostMessage(ExplorerNavigationAccess.ActiveTab(window), WinApi.WM_COMMAND, 0xA21B, 0),
                "The owned window must accept a new-tab command.");
            await WaitUntil(() => ExplorerTabAutomation.ReadSessionTabs(window).Length == count);
        }
        var tabs = ExplorerTabAutomation.ReadSessionTabs(window);
        using var watcher = new ExplorerWatcher();
        using var hook = new ExplorerTabDoubleClickHook(watcher, () => identity.IsCurrent && WinApi.GetForegroundWindow() == window);
        hook.StartHook();
        var nativeTabs = new Dictionary<string, nint>();
        foreach (var tab in tabs) await Select(tab.Id);
        await Select(tabs[1].Id);
        await Select(tabs[0].Id);
        await Select(tabs[3].Id);
        await Close(tabs[3].Id, tabs[0].Id, 3);
        // The first click on a background tab activates it; its close must still return to the old active tab.
        await Close(tabs[2].Id, tabs[0].Id, 2);
        await Close(tabs[0].Id, tabs[1].Id, 1);
        await Close(tabs[1].Id, null, 0);

        async Task Select(string id)
        {
            Check.That(ExplorerTabAutomation.TrySelectTab(window, id,
                () => identity.IsCurrent && WinApi.GetForegroundWindow() == window), "The owned tab must be selectable.");
            await WaitUntil(() => ExplorerTabAutomation.ReadSessionTabs(window).Any(tab => tab.Id == id && tab.Selected));
            nativeTabs[id] = ExplorerNavigationAccess.ActiveTab(window);
            await Task.Delay(200); // Let the real focus event reach the history observer.
        }

        async Task Close(string id, string? expected, int count)
        {
            var tab = ExplorerTabAutomation.ReadSessionTabs(window).Single(tab => tab.Id == id);
            var point = new Point((int)(tab.Bounds.Left + tab.Bounds.Width / 2), (int)(tab.Bounds.Top + tab.Bounds.Height / 2));
            await WaitUntil(() => watcher.TabStrip.IsPointOnTabStrip(point, window));
            Check.That(SetCursorPos(point.X, point.Y), "The cursor must reach the owned tab.");
            using var sampling = new CancellationTokenSource();
            var seen = new ConcurrentQueue<nint>();
            var samples = Task.Run(async () =>
            {
                while (!sampling.IsCancellationRequested)
                {
                    if (ExplorerNavigationAccess.ReadTabs(window) is { Length: > 0 } current)
                        seen.Enqueue(current[0]);
                    await Task.Delay(1);
                }
            });
            try
            {
                await WaitUntil(() => !seen.IsEmpty);
                for (var click = 0; click < 2; click++)
                {
                    Check.That(identity.IsCurrent && WinApi.GetForegroundWindow() == window &&
                        WinApi.GetAncestor(WinApi.WindowFromPoint(point), WinApi.GA_ROOT) == window,
                        "Refusing mouse input outside the owned foreground window.");
                    INPUT[] input = [Mouse(0x0002), Mouse(0x0004)];
                    Check.Equal(2u, WinApi.SendInput(2, input, Marshal.SizeOf<INPUT>()), "Both mouse events must be sent.");
                    await Task.Delay(60);
                }
                await WaitUntil(() => count == 0 ? !identity.IsCurrent :
                    ExplorerTabAutomation.ReadSessionTabs(window) is var current && current.Length == count &&
                    current.All(item => item.Id != id) && current.Any(item => item.Id == expected && item.Selected));
                if (expected != null) await WaitUntil(() => seen.Contains(nativeTabs[expected]));
            }
            finally
            {
                sampling.Cancel();
                await samples;
            }
            var allowed = expected == null ? new[] { nativeTabs[id] } : new[] { nativeTabs[id], nativeTabs[expected] };
            Check.That(seen.All(allowed.Contains),
                "The close must activate the previous tab directly, without an intermediate adjacent tab: " +
                string.Join(",", seen.Distinct().Select(value => value.ToString())));
            Console.WriteLine($"PASS Explorer double-click: remaining={count}, previous activation selected");
            // Keep separate gestures outside the controller's native double-click interval.
            await Task.Delay(100);
        }
    }

    private static async Task WaitUntil(Func<bool> condition)
    {
        var deadline = Environment.TickCount64 + 4000;
        while (!condition())
        {
            Check.That(Environment.TickCount64 < deadline, "The expected Explorer tab state was not observed.");
            await Task.Delay(20);
        }
    }

    private static INPUT Mouse(uint flags) => new()
    {
        Type = InputType.Mouse,
        Data = new InputUnion { Mouse = new MOUSEINPUT { dwFlags = flags } }
    };

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);
}
