using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.WinAPI;

/// <summary>
/// Exercises USER32's actual frame/layout cache on never-shown windows, not Explorer or the taskbar.
/// Message-only handles cannot reproduce the work-area rules for maximized top-level windows.
/// </summary>
internal static class ExplorerWindowGeometryTests
{
    private const int GwlStyle = -16;
    private const int ToolWindow = 0x80;
    private const uint OverlappedWindow = 0x00cf0000;
    private const uint Maximized = 0x01000000;
    private const uint Minimized = 0x20000000;

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("window recovery recalculates maximized work-area bounds before reveal", () => RecoverMaximized(false));
        yield return ("window recovery refreshes maximized bounds even if a prior attempt cleared the tool style", () => RecoverMaximized(true));
        yield return ("window recovery preserves normal minimized and original tool-window placement", PreservePlacement);
    }

    private static Task RecoverMaximized(bool styleAlreadyCleared) => OnWindowThread(() =>
    {
        // Both DPI contexts use USER32's own original bounds as the oracle, with no fixed taskbar height,
        // screen origin or border size. Test every attached monitor without moving any real window.
        foreach (var dpiContext in new nint[] { -4, -1 })
        {
            var previousDpi = SetThreadDpiAwarenessContext(dpiContext);
            Check.That(previousDpi != 0, "Select the native DPI context for geometry validation");
            try
            {
                foreach (var monitor in GetMonitorBounds())
                {
                    RECT expected = default;
                    WithHiddenWindow(0, Maximized, monitor, reference => expected = Bounds(reference));
                    WithHiddenWindow(ToolWindow, Maximized, monitor, handle =>
                    {
                        // Seed the exact layout produced when Explorer first maximizes a concealed
                        // preload. Creating it already maximized avoids any ShowWindow/activation.
                        var concealed = Bounds(handle);
                        if (!monitor.Work.Equals(monitor.Monitor))
                            Check.That(!expected.Equals(concealed), "Tool-window maximization must reproduce the work-area overflow");
                        WinApi.SetWindowLong(handle, WinApi.GWL_EXSTYLE,
                            WinApi.GetWindowLong(handle, WinApi.GWL_EXSTYLE) & ~ToolWindow);
                        var placement = Placement(handle);
                        var identity = WindowIdentity.Capture(handle);
                        ExplorerWindowVisibility.Hide(identity, _ => true).GetAwaiter().GetResult();
                        if (styleAlreadyCleared)
                            WinApi.SetWindowLong(handle, WinApi.GWL_EXSTYLE,
                                WinApi.GetWindowLong(handle, WinApi.GWL_EXSTYLE) & ~ToolWindow);

                        Check.That(ExplorerWindowVisibility.Restore(identity, true, _ => true), "Recover the concealed maximized window");
                        Check.That(expected.Equals(Bounds(handle)),
                            $"Recovered maximized bounds must match native work-area bounds, expected {Format(expected)}, actual {Format(Bounds(handle))}, concealed {Format(concealed)}");
                        Check.That((WinApi.GetWindowLong(handle, GwlStyle) & Maximized) != 0, "Recovery must not toggle out of maximized state");
                        Check.That(placement.Normal.Equals(Placement(handle).Normal), "Recovery must preserve the user's restore-down rectangle");
                        Check.That(!WinApi.IsWindowVisible(handle), "Recovery must never show a previously hidden frame");
                        Check.That(WinApi.GetForegroundWindow() != handle, "Geometry recovery must not activate the frame");
                    });
                }
            }
            finally { SetThreadDpiAwarenessContext(previousDpi); }
        }
    });

    private static Task PreservePlacement() => OnWindowThread(() =>
    {
        var monitor = GetMonitorBounds()[0];
        foreach (var (extendedStyle, state) in new (uint, uint)[] { (0, 0), (0, Maximized), (0, Minimized), (ToolWindow, Maximized) })
            WithHiddenWindow(extendedStyle, state, monitor, handle =>
            {
                var expected = Bounds(handle);
                var placement = Placement(handle);
                var identity = WindowIdentity.Capture(handle);
                ExplorerWindowVisibility.Hide(identity, _ => true).GetAwaiter().GetResult();
                Check.That(ExplorerWindowVisibility.Restore(identity, true, _ => true), "Recover the concealed frame");
                Check.That(expected.Equals(Bounds(handle)), "Normal, already-correct maximized, minimized and originally-tool windows must keep their bounds");
                var restored = Placement(handle);
                Check.That(placement.Normal.Equals(restored.Normal) && placement.ShowCommand == restored.ShowCommand,
                    "Frame refresh must preserve the restore-down rectangle and show state");
                Check.That(!WinApi.IsWindowVisible(handle), "Placement tests must never display a window");
                Check.That(WinApi.GetForegroundWindow() != handle, "Placement recovery must not activate the frame");
            });
    });

    private static async Task OnWindowThread(Action action)
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(action, CancellationToken.None, TaskCreationOptions.None, scheduler);
    }

    private static void WithHiddenWindow(uint extendedStyle, uint state, MonitorInfo monitor, Action<nint> action)
    {
        // Deliberately omit WS_VISIBLE and never call ShowWindow, including during cleanup.
        var handle = CreateWindowEx(extendedStyle, "STATIC", "WinTab hidden geometry test", OverlappedWindow | state,
            monitor.Work.Left + 50, monitor.Work.Top + 50, 400, 300, 0, 0, 0, 0);
        Check.That(handle != 0, "Create a private, never-shown geometry test window");
        try
        {
            Check.That(!WinApi.IsWindowVisible(handle), "The geometry fixture must start hidden");
            action(handle);
        }
        finally
        {
            ExplorerWindowVisibility.Forget(handle);
            DestroyWindow(handle);
        }
    }

    private static RECT Bounds(nint handle)
    {
        Check.That(WinApi.GetWindowRect(handle, out var rectangle), "Read native window bounds");
        return rectangle;
    }

    private static WindowPlacement Placement(nint handle)
    {
        var placement = new WindowPlacement { Length = Marshal.SizeOf<WindowPlacement>() };
        Check.That(GetWindowPlacement(handle, ref placement), "Read native restore-down placement");
        return placement;
    }

    private static string Format(RECT rectangle) => $"({rectangle.Left},{rectangle.Top},{rectangle.Right},{rectangle.Bottom})";

    private static List<MonitorInfo> GetMonitorBounds()
    {
        var monitors = new List<MonitorInfo>();
        MonitorCallback callback = (nint handle, nint dc, ref RECT rectangle, nint data) =>
        {
            var info = new MonitorInfo { Size = Marshal.SizeOf<MonitorInfo>() };
            if (GetMonitorInfo(handle, ref info)) monitors.Add(info);
            return true;
        };
        Check.That(EnumDisplayMonitors(0, 0, callback, 0) && monitors.Count > 0, "Read the attached monitor work areas");
        return monitors;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MonitorInfo { public int Size; public RECT Monitor, Work; public uint Flags; }
    [StructLayout(LayoutKind.Sequential)]
    private struct WindowPlacement
    {
        public int Length;
        public uint Flags, ShowCommand;
        public System.Drawing.Point MinPosition, MaxPosition;
        public RECT Normal;
    }
    private delegate bool MonitorCallback(nint monitor, nint dc, ref RECT rectangle, nint data);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint CreateWindowEx(uint extendedStyle, string className, string name, uint style,
        int left, int top, int width, int height, nint parent, nint menu, nint instance, nint parameter);
    [DllImport("user32.dll")]
    private static extern bool DestroyWindow(nint handle);
    [DllImport("user32.dll")]
    private static extern nint SetThreadDpiAwarenessContext(nint context);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern bool GetMonitorInfo(nint monitor, ref MonitorInfo info);
    [DllImport("user32.dll")]
    private static extern bool EnumDisplayMonitors(nint dc, nint clip, MonitorCallback callback, nint data);
    [DllImport("user32.dll")]
    private static extern bool GetWindowPlacement(nint handle, ref WindowPlacement placement);
}
