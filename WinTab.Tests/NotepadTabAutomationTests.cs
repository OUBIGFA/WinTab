using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.WinAPI;

/// <summary>
/// Exercises the Notepad tab automation against a Notepad process the test owns. Windows 11 Notepad is
/// single-instance: launching it while one is open adds a tab to the user's window, so the test skips
/// whenever a Notepad window already exists. Skips too on Windows whose Notepad has no tab strip.
/// </summary>
internal static class NotepadTabAutomationTests
{
    private const int WaitMs = 5_000;

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("notepad double-click closes tabs through the real mouse hook", ClosesTabUnderPoint);
        yield return ("notepad tab bounds exclude the space beside tab titles", IgnoresPointOutsideTabTitles);
    }

    /// <summary>Single-instance Notepad would open the temp file as a tab of the user's own window.</summary>
    private static void SkipWhenNotepadIsAlreadyRunning()
    {
        if (WinApi.FindAllWindowsEx("Notepad").Any())
            throw new TestSkippedException("A Notepad window is already open on this desktop");
    }

    private static async Task ClosesTabUnderPoint()
    {
        using var fixture = await LaunchOwnedNotepadAsync(tabs: 2);
        await RunOnStaAsync(() =>
        {
            RunGestureProbe(fixture.Window).GetAwaiter().GetResult();

            var remaining = AwaitTabs(fixture.Window, WaitMs, count => count == 1);
            Check.That(remaining is { Length: 1 }, $"One tab must remain after the close, found {remaining?.Length ?? 0}.");
            Check.Equal(fixture.FirstTabTitle, remaining![0].Title, "The tab that was not clicked must remain open.");
        });
    }

    private static async Task IgnoresPointOutsideTabTitles()
    {
        using var fixture = await LaunchOwnedNotepadAsync(tabs: 1);
        await RunOnStaAsync(() =>
        {
            var tab = fixture.Tabs[0];
            var outside = new System.Drawing.Point((int)(tab.Bounds.Right) + 40, (int)(tab.Bounds.Top + tab.Bounds.Height / 2));
            Check.That(!ExplorerTabAutomation.ReadTabs(fixture.Window).Any(tab => tab.Bounds.Contains(outside.X, outside.Y)),
                "A point beside the tab titles must not match a tab.");

            var remaining = AwaitTabs(fixture.Window, 300, count => count == 1);
            Check.That(remaining is { Length: 1 }, "The untouched tab must remain open.");
        });
    }

    private static System.Drawing.Point Center(System.Windows.Rect bounds) =>
        new((int)(bounds.X + bounds.Width / 2), (int)(bounds.Y + bounds.Height / 2));

    /// <summary>Caller supplies a disposable two-tab window; input is refused if it is covered or loses foreground.</summary>
    internal static async Task<int> RunGestureProbe(nint window)
    {
        Check.That(ExplorerWindowDiscovery.IsNotepadWindow(window), "Probe requires a Notepad window.");
        using var watcher = new ExplorerWatcher();
        using var hook = new ExplorerTabDoubleClickHook(watcher, () => WinApi.GetForegroundWindow() == window);
        hook.StatusChanged += status => Console.WriteLine(status);
        var tabs = ExplorerTabAutomation.ReadTabs(window);
        Check.Equal(2, tabs.Length, "Probe requires exactly two disposable tabs.");
        hook.StartHook();
        // No warm-up click or shared bounds-cache read: exercise the first physical double-click.
        var point = Center(tabs[1].Bounds);
        await DoubleClick(point);
        var timer = Stopwatch.StartNew();
        do
        {
            tabs = ExplorerTabAutomation.ReadTabs(window);
            if (tabs.Length == 1) break;
            await Task.Delay(10);
        } while (timer.ElapsedMilliseconds < 1000);
        Check.Equal(1, tabs.Length, "The first cold double-click must close exactly its target tab.");
        Console.WriteLine($"PASS Notepad cold double-click: observed close after {timer.ElapsedMilliseconds} ms");

        // Repeated closes must not rely on a stale rectangle or a retry click.
        for (var repeat = 0; repeat < 3; repeat++)
        {
            InvokeAddTabButton(window);
            var pair = AwaitTabs(window, WaitMs, count => count == 2)!;
            // Include an inactive tab: the first click must select it and the second close that tab.
            if (repeat == 1)
            {
                var items = AutomationElement.FromHandle(window).FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
                ((SelectionItemPattern)items[0].GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
            }
            await DoubleClick(Center(pair[1].Bounds));
            var remaining = AwaitTabs(window, 1000, count => count == 1);
            Check.That(remaining is { Length: 1 }, $"Each double-click must close one tab without retries: repeat={repeat}, count={remaining?.Length}, foreground={WinApi.GetForegroundWindow()}, expectedWindow={window}.");
        }
        Check.That(WinApi.GetWindowRect(window, out var rect), "Window must still exist.");
        await DoubleClick(new System.Drawing.Point(rect.Left + 100, rect.Top + 180));
        await Task.Delay(150);
        Check.Equal(1, ExplorerTabAutomation.ReadTabs(window).Length, "Editor double-click must not close a tab.");
        Console.WriteLine("PASS Notepad repeated closes and editor double-click");
        return 0;

        async Task DoubleClick(System.Drawing.Point target)
        {
            Helper.RestoreWindowToForeground(window);
            Check.That(SetCursorPos(target.X, target.Y), "Cursor must reach the owned window.");
            for (var click = 0; click < 2; click++)
            {
                Check.That(WinApi.GetForegroundWindow() == window &&
                    WinApi.GetAncestor(WinApi.WindowFromPoint(target), WinApi.GA_ROOT) == window,
                    "Refusing input outside the foreground owned window.");
                var input = new[] { Mouse(0x0002), Mouse(0x0004) };
                Check.Equal(2u, WinApi.SendInput(2, input, Marshal.SizeOf<INPUT>()), "Both left-button events must be delivered.");
                if (click == 0) await Task.Delay(60);
            }
        }
    }

    private static INPUT Mouse(uint flags) => new()
    {
        Type = InputType.Mouse,
        Data = new InputUnion { Mouse = new MOUSEINPUT { dwFlags = flags } }
    };

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    private static async Task<OwnedNotepad> LaunchOwnedNotepadAsync(int tabs)
    {
        SkipWhenNotepadIsAlreadyRunning();
        var directory = Path.Combine(Path.GetTempPath(), "WinTab.Tests");
        Directory.CreateDirectory(directory);
        var file = Path.Combine(directory, $"notepad-{Guid.NewGuid():N}.txt");
        await File.WriteAllTextAsync(file, "WinTab Notepad automation test");
        var process = Process.Start(new ProcessStartInfo
        {
            FileName = "notepad.exe",
            Arguments = $"\"{file}\"",
            UseShellExecute = true
        }) ?? throw new InvalidOperationException("Notepad could not be started.");

        var fixture = new OwnedNotepad(process, file);
        try
        {
            await RunOnStaAsync(() =>
            {
                fixture.Window = AwaitWindow(process, WaitMs);
                if (fixture.Window == 0)
                    throw new TestSkippedException("Notepad did not create its own window on this system");

                var initial = AwaitTabs(fixture.Window, WaitMs, count => count >= 1);
                if (initial is not { Length: >= 1 })
                    throw new TestSkippedException("This Windows' Notepad has no tab strip WinTab can read");
                fixture.FirstTabTitle = initial[0].Title;
                fixture.Tabs = initial;

                if (tabs > 1)
                {
                    InvokeAddTabButton(fixture.Window);
                    fixture.Tabs = AwaitTabs(fixture.Window, WaitMs, count => count >= tabs) ??
                        throw new InvalidOperationException($"The second Notepad tab did not appear within {WaitMs} ms");
                }
            });
            return fixture;
        }
        catch
        {
            fixture.Dispose();
            throw;
        }
    }

    private static void InvokeAddTabButton(nint window)
    {
        var root = AutomationElement.FromHandle(window);
        var button = root.FindFirst(TreeScope.Descendants,
            new PropertyCondition(AutomationElement.AutomationIdProperty, "AddButton"));
        if (button == null)
            throw new TestSkippedException("This Windows' Notepad has no new-tab button");
        ((InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
    }

    private static nint AwaitWindow(Process process, int timeoutMs)
    {
        var deadline = Environment.TickCount64 + timeoutMs;
        while (Environment.TickCount64 < deadline)
        {
            process.Refresh();
            if (process.HasExited) return 0;
            if (process.MainWindowHandle != 0)
                return process.MainWindowHandle;
            Thread.Sleep(100);
        }
        return 0;
    }

    private static ExplorerTabAutomation.Tab[]? AwaitTabs(nint window, int timeoutMs, Func<int, bool> done)
    {
        var deadline = Environment.TickCount64 + timeoutMs;
        ExplorerTabAutomation.Tab[]? last = null;
        while (Environment.TickCount64 < deadline)
        {
            last = ExplorerTabAutomation.ReadSessionTabs(window);
            if (done(last.Length))
                return last;
            Thread.Sleep(100);
        }
        return last;
    }

    private static Task RunOnStaAsync(Action action)
    {
        using var scheduler = new StaTaskScheduler();
        return Task.Factory.StartNew(action, CancellationToken.None, TaskCreationOptions.None, scheduler);
    }

    private sealed class OwnedNotepad(Process process, string file) : IDisposable
    {
        public nint Window { get; set; }
        public string FirstTabTitle { get; set; } = string.Empty;
        public ExplorerTabAutomation.Tab[] Tabs { get; set; } = [];

        public void Dispose()
        {
            try
            {
                if (!process.HasExited) process.Kill();
            }
            catch
            {
                // The process may have exited on its own.
            }
            try { File.Delete(file); } catch { }
            process.Dispose();
        }
    }
}
