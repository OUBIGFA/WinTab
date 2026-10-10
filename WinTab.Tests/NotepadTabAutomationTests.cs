using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using WinTab.Hooks;

/// <summary>Exercise the shared Notepad tab policies without launching Notepad, creating a window or sending input.</summary>
internal static class NotepadTabAutomationTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("notepad tab policy recognizes a cold double-click without native input", ColdDoubleClick);
        yield return ("notepad tab policy excludes editor and empty title-bar space", TabBoundaries);
        yield return ("notepad tab close selects the previous nonadjacent tab before closing", SelectsBeforeClosing);
        yield return ("notepad tab close stops when selection or ownership changes", InvalidatedClose);
    }

    private static ExplorerTabAutomation.Tab[] Tabs(string selected) =>
    [
        new(new System.Windows.Rect(20, 10, 140, 32), selected == "first", "same title", "first"),
        new(new System.Windows.Rect(160, 10, 140, 32), selected == "middle", "same title", "middle"),
        new(new System.Windows.Rect(300, 10, 140, 32), selected == "last", "same title", "last")
    ];

    private static Task ColdDoubleClick()
    {
        var controller = new ExplorerTabDoubleClickCloseController(new TabEnvironment { Deferred = true });
        var point = new Point(370, 26);
        Check.That(!controller.HandleLeftMouseDown(point, 1000).Handled, "The first click retains native selection");
        Check.That(!controller.HandleLeftMouseUp(1010).Handled, "The first release passes through");
        Check.That(!controller.HandleLeftMouseDown(point, 1080).Handled, "A cold hit test must not swallow unverified input");
        var decision = controller.HandleLeftMouseUp(1090);
        Check.That(!decision.Handled && decision.CloseRequest.HasValue, "A cold gesture still requests deferred tab validation");
        Check.Equal(point, decision.CloseRequest!.Value.Point);
        return Task.CompletedTask;
    }

    private static Task TabBoundaries()
    {
        foreach (var (point, onTab) in new[] { (new Point(370, 26), true),
            (new Point(480, 26), false), (new Point(100, 180), false) })
        {
            var controller = new ExplorerTabDoubleClickCloseController(new TabEnvironment());
            controller.HandleLeftMouseDown(point, 1000);
            controller.HandleLeftMouseUp(1010);
            controller.HandleLeftMouseDown(point, 1080);
            var decision = controller.HandleLeftMouseUp(1090);
            Check.Equal(onTab, decision.CloseRequest.HasValue, "Only a tab title may produce a close");
        }
        return Task.CompletedTask;
    }

    private static Task SelectsBeforeClosing()
    {
        var history = new TabActivationHistory();
        history.Observe(Tabs("middle"));
        history.Observe(Tabs("first"));
        history.Observe(Tabs("last"));
        var returnId = history.GetReturnOrder("last")[0];
        Check.Equal("first", returnId, "Activation order must take precedence over visual adjacency and duplicate titles");
        var selected = "last";
        var actions = new List<string>();
        Check.That(ExplorerTabAutomation.SelectReturnTabThenClose(() => true,
            () => { selected = returnId; actions.Add("select:" + returnId); },
            () => selected == returnId,
            () => { Check.Equal("first", selected, "The return tab must already be selected before close"); actions.Add("close:last"); }),
            "A valid transition must complete");
        Check.Equal("select:first,close:last", string.Join(',', actions));
        return Task.CompletedTask;
    }

    private static Task InvalidatedClose()
    {
        foreach (var scenario in new[] { "expired-before-selection", "selection-rejected", "expired-during-selection" })
        {
            var current = scenario != "expired-before-selection";
            var selects = 0;
            var closes = 0;
            var result = ExplorerTabAutomation.SelectReturnTabThenClose(() => current,
                () => { selects++; if (scenario == "expired-during-selection") current = false; },
                () => scenario != "selection-rejected", () => closes++);
            Check.That(!result, "A stale or rejected selection must abort: " + scenario);
            Check.Equal(0, closes, "A failed transition must never close the original tab");
            Check.Equal(scenario == "expired-before-selection" ? 0 : 1, selects);
        }
        return Task.CompletedTask;
    }

    // Only the external hit-test boundary is represented by data; gesture and close ordering are production code.
    private sealed class TabEnvironment : IExplorerTabDoubleClickEnvironment
    {
        public bool Deferred { get; init; }
        public bool IsEnabled => true;
        public int DoubleClickTimeMs => 500;
        public int DoubleClickWidth => 8;
        public int DoubleClickHeight => 8;
        public nint ResolveExplorerWindow(Point point) => 42;
        public bool IsExplorerWindow(nint window) => window == 42;
        public bool? IsPointOnTabStrip(Point point, nint window) => Deferred ? null :
            Tabs("last").Any(tab => tab.Bounds.Contains(point.X, point.Y));
    }
}
