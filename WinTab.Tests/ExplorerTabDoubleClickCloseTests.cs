using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;

internal static class ExplorerTabDoubleClickCloseTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("double-click deferred hit testing preserves cold gestures and native input", DeferredHitTestPreservesColdGesture);
        yield return ("double-click close can continue after a tab-strip hit-test refresh gap", ContinuousDoubleClicksCloseNextTabWithoutIntermediateClick);
        yield return ("double-click close chain ignores points outside the double-click geometry", CloseChainFallbackIgnoresDifferentPoints);
        yield return ("double-click close is inert while the feature is disabled", DisabledEnvironmentNeverSwallowsClicks);
        yield return ("disabling double-click close cancels an already pending close", DisablingCancelsPendingClose);
        yield return ("double-click restart discards the previous click and pending release", RestartDiscardsClickState);
        yield return ("queued double-click close rechecks ownership after its delay", QueuedCloseRechecksOwnership);
        yield return ("double-click close scope recognizes Notepad windows only when included", NotepadScopeGatesDoubleClickTargets);
        yield return ("double-click return follows activation order across tab moves and duplicate titles", ReturnFollowsActivationOrder);
        yield return ("double-click return skips closed tabs and isolates window histories", ReturnSkipsClosedTabs);
        yield return ("double-click return waits for exactly its own tab to close", ReturnRequiresConfirmedClose);
        yield return ("double-click return preserves history across incomplete snapshots", IncompleteSnapshotKeepsHistory);
    }

    private static ExplorerTabAutomation.Tab[] Tabs(string selected, params string[] ids) =>
        ids.Select(id => new ExplorerTabAutomation.Tab(new System.Windows.Rect(0, 0, 100, 30),
            id == selected, "same title", id)).ToArray();

    private static Task ReturnFollowsActivationOrder()
    {
        var history = new TabActivationHistory();
        history.Observe(Tabs("b", "a", "b", "c", "d"));
        history.Observe(Tabs("a", "a", "b", "c", "d"));
        history.Observe(Tabs("d", "a", "b", "c", "d"));
        var before = Tabs("d", "d", "c", "a", "b");
        history.Observe(before);
        Check.Equal("a", TabActivationHistory.FindReturnTab(history.GetReturnOrder("d"), "d", before,
            Tabs("c", "c", "a", "b")), "The previous activation must win over visual order and identical titles.");
        history.Observe(Tabs("a", "c", "a", "b"));
        Check.Equal("b", history.GetReturnOrder("a")[0], "A consecutive close must continue through real activations.");
        return Task.CompletedTask;
    }

    private static Task ReturnSkipsClosedTabs()
    {
        var first = new TabActivationHistory();
        var second = new TabActivationHistory();
        first.Observe(Tabs("a", "a", "b", "c"));
        first.Observe(Tabs("b", "a", "b", "c"));
        first.Observe(Tabs("c", "a", "b", "c"));
        first.Observe(Tabs("c", "a", "c"));
        second.Observe(Tabs("c", "a", "c"));
        Check.Equal("a", first.GetReturnOrder("c").Single(), "An already closed previous tab must be skipped.");
        Check.Equal(0, second.GetReturnOrder("c").Length, "Another window must not borrow activations.");
        return Task.CompletedTask;
    }

    private static Task ReturnRequiresConfirmedClose()
    {
        var before = Tabs("c", "a", "b", "c");
        foreach (var after in new[] { before, Tabs("b", "b"), Tabs("b", "a", "b", "new"), Tabs("b", "b", "new") })
            Check.That(TabActivationHistory.FindReturnTab(["a", "b"], "c", before, after) == null,
                "A save dialog or concurrent tab change must never trigger a return switch.");
        Check.Equal("a", TabActivationHistory.FindReturnTab(["a", "b"], "c", before, Tabs("b", "a", "b")),
            "Only the exact confirmed close permits selecting the previous tab.");
        return Task.CompletedTask;
    }

    private static Task IncompleteSnapshotKeepsHistory()
    {
        var history = new TabActivationHistory();
        history.Observe(Tabs("a", "a", "b"));
        history.Observe(Tabs("b", "a", "b"));
        history.Observe([]);
        history.Observe(Tabs("", "", "b"));
        Check.Equal("a", history.GetReturnOrder("b").Single(), "Incomplete UIA publication must not lose the previous tab.");
        return Task.CompletedTask;
    }

    private static Task DeferredHitTestPreservesColdGesture()
    {
        var environment = new FakeDoubleClickEnvironment { DeferHitTest = true };
        var controller = new ExplorerTabDoubleClickCloseController(environment);
        var point = new Point(240, 48);
        Check.That(!controller.HandleLeftMouseDown(point, 1000).Handled, "The first click must select normally.");
        Check.That(!controller.HandleLeftMouseUp(1020).Handled, "The first release must pass through.");
        Check.That(!controller.HandleLeftMouseDown(point, 1080).Handled, "Unverified double-clicks must remain native.");
        var up = controller.HandleLeftMouseUp(1100);
        Check.That(!up.Handled && up.CloseRequest.HasValue, "A cold double-click must reach asynchronous tab validation without swallowing native input.");
        controller.HandleLeftMouseDown(point, 1200);
        controller.HandleLeftMouseUp(1220);
        controller.HandleLeftMouseDown(point, 1280);
        environment.IsEnabled = false;
        Check.That(controller.HandleLeftMouseUp(1300).CloseRequest is null, "Disabling must cancel deferred work.");
        return Task.CompletedTask;
    }

    private static Task ContinuousDoubleClicksCloseNextTabWithoutIntermediateClick()
    {
        var environment = new FakeDoubleClickEnvironment();
        var controller = new ExplorerTabDoubleClickCloseController(environment);
        var closeRequests = new List<ExplorerTabCloseRequest>();
        var point = new Point(240, 48);

        environment.HitTestResults.Enqueue(true);
        Check.That(!controller.HandleLeftMouseDown(point, 1_000).Handled,
            "The first click should arm a tab-title candidate without swallowing native Explorer behavior.");
        Check.That(!controller.HandleLeftMouseUp(1_020).Handled,
            "A normal first mouse-up should not be swallowed.");

        environment.HitTestResults.Enqueue(true);
        Check.That(controller.HandleLeftMouseDown(point, 1_080).Handled,
            "The second click on the same tab title should be swallowed and converted into a close request.");
        RecordClose(controller.HandleLeftMouseUp(1_100), closeRequests);
        Check.Equal(1, closeRequests.Count, "The first double-click should close one tab.");

        environment.HitTestResults.Enqueue(false);
        Check.That(!controller.HandleLeftMouseDown(point, 1_220).Handled,
            "The first click after closing should still arm the next tab even while the tab-strip hit-test cache is refreshing.");
        Check.That(!controller.HandleLeftMouseUp(1_240).Handled,
            "The first mouse-up in the next pair should not be swallowed.");

        environment.HitTestResults.Enqueue(false);
        Check.That(controller.HandleLeftMouseDown(point, 1_300).Handled,
            "The second click after the refresh gap should close the next tab without requiring an intermediate click.");
        RecordClose(controller.HandleLeftMouseUp(1_320), closeRequests);
        Check.Equal(2, closeRequests.Count, "Two consecutive double-clicks should close two tabs.");

        return Task.CompletedTask;
    }

    private static Task CloseChainFallbackIgnoresDifferentPoints()
    {
        var environment = new FakeDoubleClickEnvironment();
        var controller = new ExplorerTabDoubleClickCloseController(environment);
        var closeRequests = new List<ExplorerTabCloseRequest>();
        var tabPoint = new Point(240, 48);
        var otherPoint = new Point(360, 90);

        environment.HitTestResults.Enqueue(true);
        _ = controller.HandleLeftMouseDown(tabPoint, 2_000);
        _ = controller.HandleLeftMouseUp(2_020);
        environment.HitTestResults.Enqueue(true);
        Check.That(controller.HandleLeftMouseDown(tabPoint, 2_080).Handled,
            "The setup double-click should be recognized.");
        RecordClose(controller.HandleLeftMouseUp(2_100), closeRequests);
        Check.Equal(1, closeRequests.Count, "The setup double-click should close one tab.");

        environment.HitTestResults.Enqueue(false);
        Check.That(!controller.HandleLeftMouseDown(otherPoint, 2_180).Handled,
            "A click away from the closed tab should not use the close-chain fallback.");
        Check.That(!controller.HandleLeftMouseUp(2_200).Handled,
            "The mouse-up away from the closed tab should remain native.");

        environment.HitTestResults.Enqueue(false);
        Check.That(!controller.HandleLeftMouseDown(otherPoint, 2_240).Handled,
            "A second click away from the closed tab should not close anything.");
        Check.That(!controller.HandleLeftMouseUp(2_260).Handled,
            "The second mouse-up away from the closed tab should remain native.");
        Check.Equal(1, closeRequests.Count, "Only the setup close should have been requested.");

        return Task.CompletedTask;
    }

    private static void RecordClose(MouseHookDecision decision, ICollection<ExplorerTabCloseRequest> closeRequests)
    {
        Check.That(decision.Handled, "The matching mouse-up should be swallowed.");
        var closeRequest = decision.CloseRequest;
        Check.That(closeRequest.HasValue, "The matching mouse-up should emit a native close request.");
        closeRequests.Add(closeRequest.GetValueOrDefault());
    }

    private static Task DisabledEnvironmentNeverSwallowsClicks()
    {
        var environment = new FakeDoubleClickEnvironment { IsEnabled = false };
        var controller = new ExplorerTabDoubleClickCloseController(environment);
        var point = new Point(240, 48);

        environment.HitTestResults.Enqueue(true);
        Check.That(!controller.HandleLeftMouseDown(point, 3_000).Handled, "A disabled hook must not swallow the first click.");
        Check.That(!controller.HandleLeftMouseUp(3_020).Handled, "A disabled hook must not swallow the first mouse-up.");
        environment.HitTestResults.Enqueue(true);
        Check.That(!controller.HandleLeftMouseDown(point, 3_080).Handled, "A disabled hook must not turn the second click into a close request.");
        var up = controller.HandleLeftMouseUp(3_100);
        Check.That(!up.Handled && up.CloseRequest is null, "A disabled hook must never emit a close request.");

        return Task.CompletedTask;
    }

    private static Task DisablingCancelsPendingClose()
    {
        var environment = new FakeDoubleClickEnvironment();
        var controller = new ExplorerTabDoubleClickCloseController(environment);
        var point = new Point(240, 48);
        environment.HitTestResults.Enqueue(true);
        controller.HandleLeftMouseDown(point, 1_000);
        controller.HandleLeftMouseUp(1_020);
        environment.HitTestResults.Enqueue(true);
        controller.HandleLeftMouseDown(point, 1_080);
        environment.IsEnabled = false;
        Check.That(controller.HandleLeftMouseUp(1_100).CloseRequest is null,
            "Disabling the feature must cancel the close, even after the second mouse-down.");
        return Task.CompletedTask;
    }

    private static Task RestartDiscardsClickState()
    {
        var environment = new FakeDoubleClickEnvironment();
        var controller = new ExplorerTabDoubleClickCloseController(environment);
        var point = new Point(240, 48);
        environment.HitTestResults.Enqueue(true);
        controller.HandleLeftMouseDown(point, 1_000);
        controller.HandleLeftMouseUp(1_020);
        controller.Reset();
        environment.HitTestResults.Enqueue(true);
        Check.That(!controller.HandleLeftMouseDown(point, 1_080).Handled, "A new hook lifetime must require a new double-click pair.");
        environment.HitTestResults.Enqueue(true);
        controller.HandleLeftMouseDown(point, 1_100);
        controller.Reset();
        Check.That(!controller.HandleLeftMouseUp(1_120).Handled, "A retired pending release must not close a tab in the new lifetime.");
        return Task.CompletedTask;
    }

    private static async Task QueuedCloseRechecksOwnership()
    {
        var current = true;
        var clicks = 0;
        var pending = ExplorerTabDoubleClickHook.ExecuteCloseWhenCurrentAsync(() => current, () => clicks++);
        current = false;
        Check.That(!await pending, "A disabled, moved, covered or retired request must not inject a late click.");
        Check.Equal(0, clicks, "Cancelled work must emit no input.");
        current = true;
        Check.That(await ExplorerTabDoubleClickHook.ExecuteCloseWhenCurrentAsync(() => current, () => clicks++),
            "A current close must still execute.");
        Check.Equal(1, clicks, "A current request must emit exactly one click.");
    }

    private static async Task NotepadScopeGatesDoubleClickTargets()
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(() =>
        {
            var handle = CreateNotepadClassedWindow();
            try
            {
                Check.That(ExplorerWindowDiscovery.IsNotepadWindow(handle), "A Notepad-classed window must be recognized.");
                Check.That(!ExplorerWindowDiscovery.IsFileExplorerWindow(handle),
                    "A Notepad window must never be mistaken for an Explorer frame.");
                Check.That(ExplorerWindowDiscovery.IsTabbedAppWindow(handle),
                    "The tab strip reader must accept Notepad windows.");
                Check.That(ExplorerWindowDiscovery.IsDoubleClickCloseTarget(handle, includeNotepad: true),
                    "Including Notepad must accept a Notepad window.");
                Check.That(!ExplorerWindowDiscovery.IsDoubleClickCloseTarget(handle, includeNotepad: false),
                    "The Explorer-only scope must refuse a Notepad window.");
            }
            finally { DestroyWindow(handle); }
        }, CancellationToken.None, TaskCreationOptions.None, scheduler);
    }

    private static nint CreateNotepadClassedWindow()
    {
        var windowClass = new TestWindowClass
        {
            Size = (uint)Marshal.SizeOf<TestWindowClass>(),
            Procedure = Marshal.GetFunctionPointerForDelegate(Procedure),
            Instance = GetModuleHandle(null),
            ClassName = "Notepad"
        };
        Check.That(RegisterClassEx(ref windowClass) != 0, "The Notepad test window class must register.");
        var handle = CreateWindowEx(0, "Notepad", "WinTab scope test", 0, 0, 0, 10, 10, (nint)(-3), 0, windowClass.Instance, 0);
        Check.That(handle != 0, "The Notepad-classed test window must be created.");
        return handle;
    }

    private delegate nint WindowProcedure(nint window, uint message, nint wParam, nint lParam);
    private static readonly WindowProcedure Procedure = DefWindowProc;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct TestWindowClass
    {
        public uint Size, Style;
        public nint Procedure;
        public int ClassExtra, WindowExtra;
        public nint Instance, Icon, Cursor, Background;
        public string? MenuName;
        public string ClassName;
        public nint SmallIcon;
    }

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern ushort RegisterClassEx(ref TestWindowClass windowClass);
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern nint CreateWindowEx(uint extendedStyle, string className, string name, uint style,
        int left, int top, int width, int height, nint parent, nint menu, nint instance, nint parameter);
    [DllImport("user32.dll")]
    private static extern bool DestroyWindow(nint handle);
    [DllImport("user32.dll")]
    private static extern nint DefWindowProc(nint window, uint message, nint wParam, nint lParam);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern nint GetModuleHandle(string? moduleName);

    private sealed class FakeDoubleClickEnvironment : IExplorerTabDoubleClickEnvironment
    {
        private readonly nint _explorerWindow = 42;

        public Queue<bool> HitTestResults { get; } = new();
        public bool IsEnabled { get; set; } = true;
        public bool DeferHitTest { get; set; }
        public bool ShouldDeferHitTest(nint window) => DeferHitTest;
        public int DoubleClickTimeMs => 500;
        public int DoubleClickWidth => 8;
        public int DoubleClickHeight => 8;

        public nint ResolveExplorerWindow(Point point) => _explorerWindow;
        public bool IsExplorerWindow(nint explorerWindow) => explorerWindow == _explorerWindow;
        public bool IsPointOnTabStrip(Point point, nint explorerWindow) =>
            HitTestResults.Count > 0 && HitTestResults.Dequeue();
    }
}
