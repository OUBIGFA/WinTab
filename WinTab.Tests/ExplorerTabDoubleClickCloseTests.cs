using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using WinTab.Hooks;

internal static class ExplorerTabDoubleClickCloseTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("double-click close can continue after a tab-strip hit-test refresh gap", ContinuousDoubleClicksCloseNextTabWithoutIntermediateClick);
        yield return ("double-click close chain ignores points outside the double-click geometry", CloseChainFallbackIgnoresDifferentPoints);
        yield return ("double-click close is inert while the feature is disabled", DisabledEnvironmentNeverSwallowsClicks);
        yield return ("disabling double-click close cancels an already pending close", DisablingCancelsPendingClose);
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

    private sealed class FakeDoubleClickEnvironment : IExplorerTabDoubleClickEnvironment
    {
        private readonly nint _explorerWindow = 42;

        public Queue<bool> HitTestResults { get; } = new();
        public bool IsEnabled { get; set; } = true;
        public int DoubleClickTimeMs => 500;
        public int DoubleClickWidth => 8;
        public int DoubleClickHeight => 8;

        public nint ResolveExplorerWindow(Point point) => _explorerWindow;
        public bool IsExplorerWindow(nint explorerWindow) => explorerWindow == _explorerWindow;
        public bool IsPointOnTabStrip(Point point, nint explorerWindow) =>
            HitTestResults.Count > 0 && HitTestResults.Dequeue();
    }
}
