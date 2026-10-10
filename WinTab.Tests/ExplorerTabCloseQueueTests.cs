using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using WinTab.Hooks;

/// <summary>Replay the production close queue and ordering policy without windows, UIA providers or synthesized input.</summary>
internal static class ExplorerTabCloseQueueTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("double-click close releases its queue immediately after delivery", NoPostCloseWait);
        yield return ("double-click consecutive pairs survive a pending earlier close", PendingPairsKeepOwnership);
        yield return ("double-click close waits only when the next gesture still sees the retiring tab", WaitsForLiveReplacement);
        yield return ("double-click close retries a cancelled close without a cooldown", CancelledCloseDoesNotEmbargoNextGesture);
        yield return ("double-click close cancels across delayed delivery and remote reads", RechecksOwnership);
        yield return ("double-click close provider failures do not lock later requests", FailureDoesNotPoisonQueue);
        yield return ("double-click close rejects ambiguous or non-tab live targets", RejectsUnsafeSnapshots);
        yield return ("double-click close transient publication retries remain bounded", PublicationRetriesAreBounded);
        yield return ("double-click close preserves nonadjacent return order through consecutive closes", ConsecutiveReturnOrder);
    }

    private static readonly ExplorerTabCloseRequest Request = new(42, new Point(150, 48));
    private static Task NoDelay(int _) => Task.CompletedTask;
    private static TaskCompletionSource Completion() => new(TaskCreationOptions.RunContinuationsAsynchronously);

    private static async Task NoPostCloseWait()
    {
        var environment = new TabEnvironment();
        environment.Read = count => count == 1 ? environment.Tabs :
            throw new InvalidOperationException("A delivered close must not trigger a blocking completion scan");
        var queue = new ExplorerTabCloseQueue(environment, NoDelay);
        Check.That(await queue.QueueAsync(Request, () => true, () => true), "The first close must be delivered");
        Check.Equal(1, environment.ReadCount, "A single gesture only reads its target snapshot");
        Check.That(!queue.HasPending, "Delivery, not a later publication timeout, releases the queue");
        Check.Equal("close:c", string.Join(',', environment.Actions), "No known return tab means an exact close without extra selection");
    }

    private static async Task PendingPairsKeepOwnership()
    {
        var controller = new ExplorerTabDoubleClickCloseController(new GestureEnvironment());
        var environment = new TabEnvironment();
        var entered = Completion();
        var release = Completion();
        var delays = 0;
        var queue = new ExplorerTabCloseQueue(environment, async _ =>
        {
            if (++delays != 1) return;
            entered.SetResult();
            await release.Task;
        });
        var firstRequest = Pair(controller, 1_000);
        var firstVersion = controller.InteractionVersion;
        var first = queue.QueueAsync(firstRequest, () => true, () => firstVersion == controller.InteractionVersion);
        await entered.Task.WaitAsync(TimeSpan.FromSeconds(2));
        Task<bool> second;
        try
        {
            var secondRequest = Pair(controller, 1_150);
            var secondVersion = controller.InteractionVersion;
            second = queue.QueueAsync(secondRequest, () => true, () => secondVersion == controller.InteractionVersion);
            Check.Equal(firstVersion, secondVersion, "The second pair must not cancel the earlier command");
        }
        finally { release.TrySetResult(); }
        var results = await Task.WhenAll(first, second).WaitAsync(TimeSpan.FromSeconds(2));
        Check.That(results.All(result => result), "Both complete pairs must close, even when the first was delayed");
        Check.Equal("c,b", string.Join(',', environment.Closed), "Resolve each live replacement in FIFO order");
        Check.Equal(2, environment.ReadCount, "Neither close needs a post-delivery scan");
    }

    private static ExplorerTabCloseRequest Pair(ExplorerTabDoubleClickCloseController controller, long time)
    {
        controller.HandleLeftMouseDown(Request.Point, time);
        controller.HandleLeftMouseUp(time + 10);
        controller.HandleLeftMouseDown(Request.Point, time + 60);
        var result = controller.HandleLeftMouseUp(time + 70);
        Check.That(result.CloseRequest.HasValue, "A complete pair must enqueue a close");
        return result.CloseRequest!.Value;
    }

    private static async Task WaitsForLiveReplacement()
    {
        var environment = new TabEnvironment { RemoveOnClose = false };
        var waits = new List<int>();
        var queue = new ExplorerTabCloseQueue(environment, milliseconds => { waits.Add(milliseconds); return Task.CompletedTask; });
        Check.That(await queue.QueueAsync(Request, () => true, () => true), "A delivered command must finish even before publication");
        Check.Equal(1, environment.ReadCount, "Do not poll after the first delivery");
        environment.RemoveOnClose = true;
        environment.Read = count =>
        {
            if (count == 2) return [];
            if (count == 3) return environment.Tabs;
            environment.Remove("c");
            return environment.Tabs;
        };
        Check.That(await queue.QueueAsync(Request, () => true, () => true), "The next pair must survive a temporarily unpublished strip");
        Check.Equal("c,b", string.Join(',', environment.Closed), "Never invoke the retiring tab twice");
        Check.Equal("10,10,15,15", string.Join(',', waits), "Wait only for actual missing/retiring targets, never a fixed cooldown");
    }

    private static async Task CancelledCloseDoesNotEmbargoNextGesture()
    {
        var environment = new TabEnvironment { RemoveOnClose = false };
        var current = true;
        var queue = new ExplorerTabCloseQueue(environment, NoDelay);
        Check.That(await queue.QueueAsync(Request, () => true, () => current), "The first close can open a save dialog");
        current = false; // The dialog or other input ends that gesture; the user can cancel it and try again.
        environment.RemoveOnClose = true;
        Check.That(await queue.QueueAsync(Request, () => true, () => true), "A new gesture must be allowed to retry the still-open tab");
        Check.Equal("c,c", string.Join(',', environment.Closed), "Each explicit gesture invokes the requested tab exactly once");
    }

    private static async Task RechecksOwnership()
    {
        var current = true;
        var environment = new TabEnvironment();
        var queue = new ExplorerTabCloseQueue(environment, _ => { current = false; return Task.CompletedTask; });
        Check.That(!await queue.QueueAsync(Request, () => current, () => current), "Ownership lost during delivery delay must abort");
        Check.Equal(0, environment.ReadCount, "A retired request does not even read tabs");

        current = true;
        environment.Read = _ => { current = false; return environment.Tabs; };
        queue = new ExplorerTabCloseQueue(environment, NoDelay);
        Check.That(!await queue.QueueAsync(Request, () => current, () => current), "Recheck after a remote read returns");
        Check.Equal(0, environment.Actions.Count, "A stale snapshot must not cause a selection or close");

        current = true;
        environment.Read = null;
        environment.PrimeHistory("d", "a", "c");
        environment.AfterSelection = () => current = false;
        Check.That(!await queue.QueueAsync(Request, () => current, () => current), "Recheck across remote selection too");
        Check.Equal(0, environment.Closed.Count, "Losing ownership during selection must not close the tab");
    }

    private static async Task FailureDoesNotPoisonQueue()
    {
        var environment = new TabEnvironment();
        environment.Read = count => count == 1 ? throw new InvalidOperationException("provider unavailable") : environment.Tabs;
        var queue = new ExplorerTabCloseQueue(environment, NoDelay);
        var first = queue.QueueAsync(Request, () => true, () => true);
        var second = queue.QueueAsync(Request, () => true, () => true);
        var failed = false;
        try { await first; }
        catch (InvalidOperationException exception) when (exception.Message == "provider unavailable") { failed = true; }
        Check.That(failed, "Do not hide the provider exception");
        Check.That(await second.WaitAsync(TimeSpan.FromSeconds(2)), "A failed predecessor must not poison the FIFO tail");
        Check.Equal(1, environment.Closed.Count, "Only the valid second request may close");
        Check.That(!queue.HasPending, "Failures must release pending accounting");
    }

    private static async Task RejectsUnsafeSnapshots()
    {
        var valid = new TabEnvironment().Tabs;
        var badSnapshots = new[]
        {
            valid.Select(tab => tab with { Id = "same" }).ToArray(),
            valid.Select(tab => tab with { Id = "" }).ToArray(),
            valid.Select(tab => tab with { Bounds = valid[0].Bounds }).ToArray(),
            valid.Select(tab => tab with { Bounds = new System.Windows.Rect(500, 30, 100, 30) }).ToArray()
        };
        foreach (var tabs in badSnapshots)
        {
            var environment = new TabEnvironment { Tabs = tabs };
            var queue = new ExplorerTabCloseQueue(environment, NoDelay);
            Check.That(!await queue.QueueAsync(Request, () => true, () => true), "Unknown identity, overlapping hits and non-tab points must be refused");
            Check.Equal(0, environment.Actions.Count, "A cached hit alone never authorizes a close");
        }
    }

    private static async Task PublicationRetriesAreBounded()
    {
        var environment = new TabEnvironment();
        environment.Read = _ => [];
        var now = 0;
        var queue = new ExplorerTabCloseQueue(environment, milliseconds => { now += milliseconds; return Task.CompletedTask; });
        Check.That(!await queue.QueueAsync(Request, () => now < 80, () => now < 80), "Missing publication must stop at the request deadline");
        Check.That(!queue.HasPending && environment.ReadCount < 10, "Do not leave a busy worker or unbounded retries");
        environment.Read = null;
        now = 0;
        Check.That(await queue.QueueAsync(Request, () => now < 80, () => now < 80), "A timed-out request must not impose a cooldown on the next one");
    }

    private static async Task ConsecutiveReturnOrder()
    {
        var environment = new TabEnvironment();
        environment.PrimeHistory("d", "a", "c");
        var queue = new ExplorerTabCloseQueue(environment, NoDelay);
        Check.That(await queue.QueueAsync(Request, () => true, () => true), "Close the first tab through its nonadjacent MRU successor");
        environment.Select("b"); // The first native click of the next pair selects the replacement title.
        Check.That(await queue.QueueAsync(Request, () => true, () => true), "The next close follows the same ordered policy");
        Check.Equal("select:a,close:c,select:a,close:b", string.Join(',', environment.Actions),
            "Only the intended successor may be selected, and always before closing the original tab");
    }

    private sealed class TabEnvironment : IExplorerTabCloseEnvironment
    {
        private readonly TabActivationHistory _history = new();
        public ExplorerTabAutomation.Tab[] Tabs { get; set; } =
            new[] { "c", "b", "a", "d" }.Select((id, index) =>
                new ExplorerTabAutomation.Tab(new System.Windows.Rect(100 + index * 120, 30, 100, 30), id == "c", "same title", id)).ToArray();
        public List<string> Actions { get; } = [];
        public List<string> Closed { get; } = [];
        public int ReadCount { get; private set; }
        public bool RemoveOnClose { get; set; } = true;
        public Func<int, ExplorerTabAutomation.Tab[]>? Read { get; set; }
        public Action? AfterSelection { get; set; }

        public void Select(string id) => Tabs = Tabs.Select(tab => tab with { Selected = tab.Id == id }).ToArray();
        public void Remove(string id) => Tabs = Tabs.Where(tab => tab.Id != id).Select((tab, index) =>
            tab with { Bounds = new System.Windows.Rect(100 + index * 120, 30, 100, 30) }).ToArray();
        public void PrimeHistory(params string[] ids)
        {
            foreach (var id in ids) { Select(id); _history.Observe(Tabs); }
        }
        public ExplorerTabAutomation.Tab[] ReadTabs(nint window)
        {
            ReadCount++;
            return Read?.Invoke(ReadCount) ?? Tabs;
        }
        public string? GetReturnTab(nint window, ExplorerTabAutomation.Tab[] tabs, string closingId)
        {
            _history.Observe(tabs);
            return _history.GetReturnOrder(closingId).FirstOrDefault();
        }
        public bool TryCloseTab(nint window, string closingId, string? returnId, Func<bool> isCurrent)
        {
            void Close()
            {
                Actions.Add("close:" + closingId);
                Closed.Add(closingId);
                if (RemoveOnClose) Remove(closingId);
            }
            if (returnId == null)
            {
                if (!isCurrent()) return false;
                Close();
                return true;
            }
            return ExplorerTabAutomation.SelectReturnTabThenClose(isCurrent,
                () => { Actions.Add("select:" + returnId); Select(returnId); AfterSelection?.Invoke(); },
                () => Tabs.Any(tab => tab.Id == returnId && tab.Selected), Close);
        }
    }

    private sealed class GestureEnvironment : IExplorerTabDoubleClickEnvironment
    {
        public bool IsEnabled => true;
        public int DoubleClickTimeMs => 500;
        public int DoubleClickWidth => 8;
        public int DoubleClickHeight => 8;
        public nint ResolveExplorerWindow(Point point) => 42;
        public bool IsExplorerWindow(nint window) => window == 42;
        public bool? IsPointOnTabStrip(Point point, nint window) => true;
    }
}
