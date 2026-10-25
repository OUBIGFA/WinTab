using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.WinAPI;

/// <summary>Exercise the production gesture, hit policy and async pipeline without desktop windows or input.</summary>
internal static class ExplorerFileListDoubleClickTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("file-list double-click requires two complete matching clicks", PairsMatchingClicks);
        yield return ("file-list double-click rejects time distance location identity and bounds changes", RejectsMismatches);
        yield return ("file-list double-click rejects drags holds and repeated downs", RejectsIncompleteClicks);
        yield return ("file-list double-click never reuses a completed pair", DoesNotReusePair);
        yield return ("file-list hit policy accepts only positively identified view backgrounds", BackgroundPolicy);
        yield return ("file-list hit policy rejects files folders text headers and interactive controls", OccupiedPolicy);
        yield return ("file-list navigation uses distinct same-tab parent and back flags", BrowseFlags);
        yield return ("file-list snapshots require a live active non-Control-Panel tab without enabling merging", LocationSnapshots);
        yield return ("file-list double-click navigates once in the selected mode", NavigatesOnce);
        yield return ("file-list double-click requires both backgrounds to be empty", RequiresTwoEmptyHits);
        yield return ("file-list double-click cancels travel away and back between or during clicks", CancelsTravel);
        yield return ("file-list double-click ignores delayed modified and self-injected input", RejectsInvalidInput);
        yield return ("file-list double-click cancels pending work on other input disable or mode change", CancelsPendingWork);
        yield return ("file-list double-click does not block input or queue unbounded UIA workers", BoundsSlowWorkers);
        yield return ("file-list double-click expires a slow hit even without further input", ExpiresWithoutFurtherInput);
        yield return ("file-list double-click contains hit-test failures and accepts a later gesture", RecoversAfterHitFailure);
    }

    private static ExplorerFileListDoubleClickState NewState() => new(() => 500, () => 4, () => 4);
    private static readonly Point Origin = new(100, 100);
    private static ExplorerFileListClick Click(long tick, Point? point = null) => new(
        new WindowIdentity(1, 10, 20, 30), new WindowIdentity(2, 10, 20, 31),
        new WindowIdentity(3, 10, 20, 32), point ?? Origin, @"C:\Parent\Child", tick)
        { ViewBounds = new RECT { Left = 0, Top = 0, Right = 500, Bottom = 500 } };

    private static Task PairsMatchingClicks()
    {
        var state = NewState();
        var first = Click(100);
        var second = Click(300, new Point(101, 101));
        Check.That(ReferenceEquals(first, state.HandleLeftDown(first)), "First down starts a candidate");
        Check.That(state.HandleLeftUp(first.Point, 110) == null, "First up performs no navigation");
        Check.That(state.HandleLeftDown(second) == null, "Second down only pairs the candidate");
        var request = state.HandleLeftUp(second.Point, 310);
        Check.That(request != null && ReferenceEquals(first, request.First) && ReferenceEquals(second, request.Second),
            "Second up returns both original snapshots");
        return Task.CompletedTask;
    }

    private static Task RejectsMismatches()
    {
        var next = Click(200);
        foreach (var second in new[] { Click(601), Click(50), Click(200, new Point(110, 100)),
            next with { Location = @"C:\Other" }, next with { Window = new WindowIdentity(9, 10, 20, 30) },
            next with { Tab = new WindowIdentity(9, 10, 20, 31) }, next with { View = new WindowIdentity(9, 10, 20, 32) },
            next with { Tab = new WindowIdentity(2, 10, 20, 99) },
            next with { ViewBounds = new RECT { Left = 10, Top = 0, Right = 510, Bottom = 500 } } })
        {
            var state = NewState();
            state.HandleLeftDown(Click(100));
            state.HandleLeftUp(Origin, 110);
            state.HandleLeftDown(second);
            Check.That(state.HandleLeftUp(second.Point, second.Tick + 1) == null, "A mismatched second click must not navigate");
        }
        return Task.CompletedTask;
    }

    private static Task RejectsIncompleteClicks()
    {
        foreach (var scenario in new[] { "drag", "hold", "repeated-down", "invalid-target" })
        {
            var state = NewState();
            state.HandleLeftDown(Click(100));
            if (scenario == "repeated-down") state.HandleLeftDown(Click(105));
            if (scenario == "invalid-target") state.HandleLeftDown(null);
            state.HandleLeftUp(scenario == "drag" ? new Point(200, 200) : Origin, scenario == "hold" ? 601 : 110);
            state.HandleLeftDown(Click(scenario == "hold" ? 700 : 200));
            Check.That(state.HandleLeftUp(Origin, scenario == "hold" ? 710 : 210) == null, "Incomplete first click cannot form a pair: " + scenario);
        }
        return Task.CompletedTask;
    }

    private static Task DoesNotReusePair()
    {
        var state = NewState();
        for (var i = 0; i < 4; i++)
        {
            state.HandleLeftDown(Click(100 + 50 * i));
            Check.Equal(i % 2 == 1, state.HandleLeftUp(Origin, 110 + 50 * i) != null,
                "Four clicks must form two non-overlapping pairs, never three");
        }
        return Task.CompletedTask;
    }

    private static Task BackgroundPolicy()
    {
        foreach (var background in new[] { "UIItemsView", "UIGroupItem" })
            Check.That(ExplorerFileListAutomation.IsBackgroundPath(
                [(ControlType.List, background, 0), (ControlType.Pane, "DirectUIHWND", 3)], 3), "Known file-list background is valid");
        Check.That(!ExplorerFileListAutomation.IsBackgroundPath([(ControlType.Pane, "DirectUIHWND", 3)], 3),
            "A generic container is not evidence of blank file-list space");
        Check.That(!ExplorerFileListAutomation.IsBackgroundPath([(ControlType.List, "UIItemsView", 9)], 3),
            "A background in a different view must be rejected");
        Check.That(!ExplorerFileListAutomation.IsBackgroundPath([(ControlType.List, "UIItemsView", 0)], 0),
            "A missing native view must be rejected");
        var deep = Enumerable.Repeat((ControlType.Pane, "unknown", (nint)0), 64)
            .Append((ControlType.List, "UIItemsView", (nint)3));
        Check.That(!ExplorerFileListAutomation.IsBackgroundPath(deep, 3), "An excessively deep provider tree must stop");
        return Task.CompletedTask;
    }

    private static Task OccupiedPolicy()
    {
        foreach (var type in new[] { ControlType.ListItem, ControlType.DataItem, ControlType.TreeItem, ControlType.Button,
            ControlType.Hyperlink, ControlType.Edit, ControlType.ComboBox, ControlType.CheckBox, ControlType.ScrollBar,
            ControlType.Thumb, ControlType.Header, ControlType.HeaderItem, ControlType.Menu, ControlType.MenuItem, ControlType.Text })
            Check.That(!ExplorerFileListAutomation.IsBackgroundPath(
                [(type, "", 0), (ControlType.List, "UIItemsView", 3)], 3), "Occupied control must veto its background ancestor: " + type.ProgrammaticName);
        return Task.CompletedTask;
    }

    private static Task BrowseFlags()
    {
        Check.Equal(0x2001u, ExplorerWatcher.FileListBrowseFlags(FileListDoubleClickAction.Parent), "Parent uses SBSP_PARENT | SBSP_SAMEBROWSER");
        Check.Equal(0x4001u, ExplorerWatcher.FileListBrowseFlags(FileListDoubleClickAction.Back), "Back uses SBSP_NAVIGATEBACK | SBSP_SAMEBROWSER");
        Check.Throws<ArgumentOutOfRangeException>(() => ExplorerWatcher.FileListBrowseFlags((FileListDoubleClickAction)99),
            "An invalid action must not turn into an unintended navigation");
        return Task.CompletedTask;
    }

    private static Task LocationSnapshots() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        // The fixture owns message-only HWNDs and a Shell stub that fails if this input-thread read touches COM.
        fixture.Window.SetActive(0);
        fixture.AddBrowser(out var info, fixture.Window.FirstTab,
            onCall: _ => throw new InvalidOperationException("A file-list snapshot must not call the Shell"));
        var window = info.Identity;
        var tab = info.TabIdentity;
        Check.That(!fixture.Watcher.IsHookActive, "Navigation metadata must remain usable while merging is off");
        info.Location = "file:///C:/Parent/Child";
        Check.Equal(@"C:\Parent\Child", fixture.Watcher.TryGetFileListLocationSnapshot(window, tab),
            "An active tracked tab supplies its normalized location without a Shell call");
        fixture.Window.SetActive(1);
        Check.That(fixture.Watcher.TryGetFileListLocationSnapshot(window, tab) == null,
            "A background tab must not supply a gesture target");
        fixture.Window.SetActive(0);
        foreach (var location in new[] { "", " ", "shell:ControlPanelFolder",
            "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\0",
            "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}" })
        {
            info.Location = location;
            Check.That(fixture.Watcher.TryGetFileListLocationSnapshot(window, tab) == null,
                "Unknown locations and Control Panel stay outside file-list navigation");
        }
        info.Location = @"C:\Parent\Child";
        info.Closed = true;
        Check.That(fixture.Watcher.TryGetFileListLocationSnapshot(window, tab) == null,
            "A closed tab must not be captured");
        info.Closed = false;
        tab.Release();
        info.TabIdentity = WindowIdentity.Capture(tab.Handle);
        Check.That(fixture.Watcher.TryGetFileListLocationSnapshot(window, tab) == null,
            "A reused HWND cannot inherit an old gesture's identity");
        return Task.CompletedTask;
    }, tabCount: 2);

    private sealed class NoInput : IDisposable { public void Dispose() { } }
    private sealed class Fixture : IDisposable
    {
        internal readonly ExplorerFileListDoubleClickHook Hook;
        internal Func<ExplorerFileListClick, Task<bool>> Hit = _ => Task.FromResult(true);
        internal Func<Point, long, ExplorerFileListClick?> Capture = (point, tick) => Click(tick, point);
        internal FileListDoubleClickAction Action, NavigatedAction;
        internal bool Valid = true, Modifiers;
        internal int Navigations, HitReads;

        internal Fixture()
        {
            Hook = new ExplorerFileListDoubleClickHook(_ => new NoInput(), (point, tick) => Capture(point, tick),
                click => { Interlocked.Increment(ref HitReads); return Hit(click); },
                (_, action, current) =>
                {
                    Check.That(current(), "The final navigation boundary must still own this gesture");
                    NavigatedAction = action;
                    Interlocked.Increment(ref Navigations);
                    return Task.FromResult(true);
                }, _ => Valid, () => Action, () => Modifiers, NewState());
            Hook.StartHook();
        }

        internal Task Input(NavigationPointerKind kind, Point? point = null, int delay = 0, bool own = false) =>
            Hook.ProcessPointerAsync(new NavigationPointerInput(kind, point ?? Origin, own, delay));
        internal async Task ClickOnce()
        {
            await Input(NavigationPointerKind.LeftDown);
            await Input(NavigationPointerKind.LeftUp);
        }
        internal async Task DoubleClick() { await ClickOnce(); await ClickOnce(); }
        public void Dispose() => Hook.Dispose();
    }

    private static async Task NavigatesOnce()
    {
        foreach (var action in Enum.GetValues<FileListDoubleClickAction>())
        {
            using var fixture = new Fixture { Action = action };
            await fixture.DoubleClick();
            Check.Equal(1, fixture.Navigations, "One completed pair produces one navigation");
            Check.Equal(action, fixture.NavigatedAction, "The selected mode reaches navigation unchanged");
            await fixture.ClickOnce();
            Check.Equal(1, fixture.Navigations, "A third click does not reuse the previous pair");
        }
    }

    private static async Task RequiresTwoEmptyHits()
    {
        for (var occupied = 1; occupied <= 2; occupied++)
        {
            using var fixture = new Fixture();
            var reads = 0;
            fixture.Hit = _ => Task.FromResult(Interlocked.Increment(ref reads) != occupied);
            await fixture.DoubleClick();
            Check.Equal(0, fixture.Navigations, "An occupied hit must veto the pair, regardless of query order");
        }
    }

    private static async Task CancelsTravel()
    {
        foreach (var between in new[] { false, true })
        {
            using var fixture = new Fixture();
            await fixture.Input(NavigationPointerKind.LeftDown);
            if (between) await fixture.Input(NavigationPointerKind.LeftUp);
            await fixture.Input(NavigationPointerKind.Move, new Point(200, 200));
            await fixture.Input(NavigationPointerKind.Move);
            if (!between) await fixture.Input(NavigationPointerKind.LeftUp);
            await fixture.ClickOnce();
            Check.Equal(0, fixture.Navigations, "Travel away and back must cancel, even when both endpoints match");
        }
    }

    private static async Task RejectsInvalidInput()
    {
        foreach (var scenario in new[] { "disabled", "modifier", "delayed", "injected", "outside" })
        {
            using var fixture = new Fixture();
            if (scenario == "disabled") fixture.Hook.StopHook();
            if (scenario == "modifier") fixture.Modifiers = true;
            if (scenario == "outside") fixture.Capture = (_, _) => null;
            for (var i = 0; i < 2; i++)
            {
                await fixture.Input(NavigationPointerKind.LeftDown, delay: scenario == "delayed" ? 101 : 0, own: scenario == "injected");
                await fixture.Input(NavigationPointerKind.LeftUp);
            }
            Check.Equal(0, fixture.Navigations, "Invalid gesture must not navigate: " + scenario);
            Check.Equal(0, fixture.HitReads, "Invalid gesture must not query UIA: " + scenario);
        }
    }

    private static async Task CancelsPendingWork()
    {
        foreach (var scenario in new[] { "keyboard", "stop", "mode", "stale-view", "wheel", "middle", "other", "new-click", "move" })
        {
            using var fixture = new Fixture();
            var release = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            fixture.Hit = _ => release.Task;
            var pending = fixture.DoubleClick();
            switch (scenario)
            {
                case "keyboard": fixture.Hook.CancelPending(); break;
                case "stop": fixture.Hook.StopHook(); fixture.Hook.StartHook(); break;
                case "mode": fixture.Action = FileListDoubleClickAction.Back; break;
                case "stale-view": fixture.Valid = false; break;
                case "wheel": await fixture.Input(NavigationPointerKind.Wheel); break;
                case "middle": await fixture.Input(NavigationPointerKind.MiddleDown); break;
                case "other": await fixture.Input(NavigationPointerKind.OtherDown); break;
                case "new-click": await fixture.Input(NavigationPointerKind.LeftDown); break;
                case "move": await fixture.Input(NavigationPointerKind.Move, new Point(200, 200)); break;
            }
            release.SetResult(true);
            await pending.WaitAsync(TimeSpan.FromSeconds(3));
            Check.Equal(0, fixture.Navigations, "Late hit results must not resurrect a cancelled gesture: " + scenario);
        }
    }

    private static async Task BoundsSlowWorkers()
    {
        using var fixture = new Fixture();
        var release = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var workers = 0;
        fixture.Hit = _ => { if (Interlocked.Increment(ref workers) == 2) entered.SetResult(); return release.Task; };
        var pending = fixture.DoubleClick();
        try
        {
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(3));
            for (var i = 0; i < 20; i++) await fixture.DoubleClick().WaitAsync(TimeSpan.FromSeconds(1));
            Check.Equal(2, fixture.HitReads, "Stalled UIA queries must not accumulate workers or a queued backlog");
            await pending.WaitAsync(TimeSpan.FromSeconds(3));
            Check.Equal(0, fixture.Navigations, "The request must retire without waiting indefinitely for its provider");
        }
        finally { release.TrySetResult(true); }
    }

    private static async Task ExpiresWithoutFurtherInput()
    {
        using var fixture = new Fixture();
        var release = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var reads = 0;
        fixture.Hit = _ => { if (Interlocked.Increment(ref reads) == 2) entered.SetResult(); return release.Task; };
        var pending = fixture.DoubleClick();
        try
        {
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(3));
            // No cancellation or newer gesture is sent: the request must expire on its own deadline.
            await pending.WaitAsync(TimeSpan.FromSeconds(3));
            Check.That(!release.Task.IsCompleted, "The input pipeline must retire while the provider is still stalled");
            Check.Equal(0, fixture.Navigations, "An expired gesture must never reach Shell navigation");
        }
        finally { release.TrySetResult(true); }
    }

    private static async Task RecoversAfterHitFailure()
    {
        using var fixture = new Fixture { Hit = _ => throw new InvalidOperationException("unavailable provider") };
        await fixture.DoubleClick();
        Check.Equal(0, fixture.Navigations, "Provider exceptions must fail closed");
        fixture.Hit = _ => Task.FromResult(true);
        await fixture.DoubleClick();
        Check.Equal(1, fixture.Navigations, "A failed provider call must release its worker slot for the next gesture");
    }
}
