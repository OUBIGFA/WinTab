using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.Managers;
using WinTab.Models;

internal static class HookManagerTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("a hook that cannot start is reported without stopping the caller", HookStartFailureIsContained);
        yield return ("a hook that cannot stop is reported without stopping the caller", HookStopFailureIsContained);
        yield return ("hook changes are applied only when the state differs", HookChangeIsIdempotent);
        yield return ("features whose hook differs from the setting are listed in order", ListsMismatchedFeatures);
        yield return ("session shortcuts apply independently of automatic restore and the history preference", IndependentSessionShortcuts);
        yield return ("session shortcuts alone record closed tabs and retain history until both recording options are off", ShortcutOnlyHistory);
    }

    private static Task ListsMismatchedFeatures()
    {
        var mismatched = HookManager.FindMismatched(
        [
            (HookFeature.MergeWindows, true, new FakeHook { IsHookActive = true }),
            (HookFeature.DoubleClickClose, true, new FakeHook()),
            (HookFeature.MiddleClickForeground, false, new FakeHook()),
            (HookFeature.WheelSwitch, false, new FakeHook { IsHookActive = true })
        ]);
        Check.Equal("DoubleClickClose,WheelSwitch", string.Join(",", mismatched),
            "An enabled hook that did not start and a disabled hook that still runs must both be listed.");
        return Task.CompletedTask;
    }

    private static Task HookStartFailureIsContained()
    {
        var hook = new FakeHook { FailStart = true };
        Check.That(!HookManager.TryChangeHookStatus(hook, true), "A failed start must be reported as a failure.");
        Check.That(!hook.IsHookActive, "A failed start must leave the hook inactive.");
        return Task.CompletedTask;
    }

    private static Task HookStopFailureIsContained()
    {
        var hook = new FakeHook { IsHookActive = true, FailStop = true };
        Check.That(!HookManager.TryChangeHookStatus(hook, false), "A failed stop must be reported as a failure.");
        return Task.CompletedTask;
    }

    private static Task HookChangeIsIdempotent()
    {
        var hook = new FakeHook();
        Check.That(HookManager.TryChangeHookStatus(hook, true) && hook.IsHookActive, "A start must activate the hook.");
        Check.That(HookManager.TryChangeHookStatus(hook, true), "Starting an active hook must succeed without starting it again.");
        Check.Equal(1, hook.Starts, "An active hook must not be started twice.");
        Check.That(HookManager.TryChangeHookStatus(hook, false) && !hook.IsHookActive, "A stop must deactivate the hook.");
        return Task.CompletedTask;
    }

    private static Task IndependentSessionShortcuts() => WithSessionShortcuts((fixture, manager, dispatch) =>
    {
        foreach (var automatic in new[] { false, true })
        foreach (var keepHistory in new[] { false, true })
        foreach (var groupEnabled in new[] { false, true })
        foreach (var tabEnabled in new[] { false, true })
        {
            var settings = new AppSettings
            {
                RestoreTabs = automatic, ReopenClosedTab = keepHistory,
                RestoreGroupShortcutEnabled = groupEnabled, ReopenTabShortcutEnabled = tabEnabled
            };
            Set(fixture.Watcher, "_restoreTabs", automatic);
            manager.ApplySessionShortcuts(settings);
            Check.That(manager.ShortcutError == null, "The independent shortcuts must configure successfully.");
            Check.Equal(keepHistory || tabEnabled, Get<bool>(fixture.Watcher, "_recordClosedTabs"),
                "Either the history preference or the tab shortcut keeps recording on.");
            Check.Equal(automatic, Get<bool>(fixture.Watcher, "_restoreTabs"), "Configuring shortcuts must not change automatic restore.");
            Check.That(!fixture.Watcher.IsHookActive, "Manual recovery must not enable window merging.");

            Check.Equal(groupEnabled, dispatch.Handle('E', true, ShortcutModifiers.Alt, false, false, out var group),
                "The group shortcut works outside Explorer even with automatic restore off.");
            Check.That(group == (groupEnabled ? SessionAction.RestoreGroup : null), "Only the group shortcut's own toggle gates its action.");
            dispatch.Handle('E', false, ShortcutModifiers.None, false, false, out _);
            Check.Equal(tabEnabled, dispatch.Handle('W', true, ShortcutModifiers.Alt, true, false, out var tab),
                "The tab shortcut does not require the separate history toggle.");
            Check.That(tab == (tabEnabled ? SessionAction.ReopenTab : null), "Only the tab shortcut's own toggle gates its action.");
            dispatch.Handle('W', false, ShortcutModifiers.None, true, false, out _);
            Check.That(!dispatch.Handle('W', true, ShortcutModifiers.Alt, false, false, out var otherApp) && otherApp == null,
                "Independent activation must not expand closed-tab recovery to browsers or other apps.");
            dispatch.Handle('W', false, ShortcutModifiers.None, false, false, out _);
        }
        return Task.CompletedTask;
    });

    private static Task ShortcutOnlyHistory() => WithSessionShortcuts(async (fixture, manager, _) =>
    {
        var settings = new AppSettings
        {
            RestoreTabs = false, ReopenClosedTab = false,
            RestoreGroupShortcutEnabled = false, ReopenTabShortcutEnabled = true
        };
        manager.ApplySessionShortcuts(settings);
        Set(fixture.Watcher, "_captureSessions", true);
        var frame = WindowIdentity.Capture(fixture.Window.Handle);
        Get<ConcurrentDictionary<nint, WindowIdentity>>(fixture.Watcher, "_sessionWindows")[frame.Handle] = frame;
        var closed = new WindowInfo
        {
            Identity = frame, TabIdentity = new WindowIdentity(0, 1, 1, 101), Location = @"C:\shortcut-only"
        };
        fixture.Invoke("RecordClosedTab", closed);
        var history = Get<ExplorerClosedTabHistory>(fixture.Watcher, "_closedTabs");
        Check.Equal(1, history.Count, "An enabled shortcut must record closes even when the separate history preference is off.");
        Check.Equal(closed.Location, history.Peek()[0].Location, "A shortcut alone must actually capture a user close.");

        // Exercise the command's history gate without opening a real Explorer window or sending native keys.
        Set(fixture.Watcher, "_sessionLocationPolicy", new ExplorerSessionLocationPolicy(_ => false));
        var result = await (Task<SessionCommandResult>)fixture.Invoke("ReopenClosedTabCoreAsync", CancellationToken.None)!;
        Check.Equal(SessionCommandResult.NothingAvailable, result, "The command reaches location validation instead of rejecting disabled recording.");
        Check.Equal(1, history.Count, "A temporarily unavailable close remains recorded.");

        manager.ApplySessionShortcuts(settings with { ReopenClosedTab = true });
        manager.ApplySessionShortcuts(settings);
        Check.Equal(1, history.Count, "Turning the separate recording preference off cannot erase history needed by the shortcut.");
        manager.ApplySessionShortcuts(settings with { ReopenClosedTab = true, ReopenTabShortcutEnabled = false });
        Check.Equal(1, history.Count, "Turning the shortcut off retains history when button-based recovery is still enabled.");
        var pending = history.Peek()[0];
        Check.That(history.TryTake(pending), "A reopen may already be in flight when the last recording option is disabled.");
        manager.ApplySessionShortcuts(settings with { ReopenTabShortcutEnabled = false });
        history.Return(pending);
        fixture.Invoke("RecordClosedTab", closed);
        Check.Equal(0, history.Count, "Opting out of both recording and the shortcut clears history and blocks late returns or new closes.");
        manager.ApplySessionShortcuts(settings);
        fixture.Invoke("RecordClosedTab", closed);
        Check.Equal(1, history.Count, "Re-enabling just the shortcut resumes recording.");
    });

    private static Task WithSessionShortcuts(Func<ExplorerTabLifetimeTests.Fixture, HookManager, ExplorerShortcutDispatch, Task> body) =>
        ExplorerTabLifetimeTests.WithFixture(async fixture =>
        {
            // Exercise configuration and dispatch without installing a process-wide keyboard hook.
            using var shortcuts = new ShortcutConfiguration();
            var manager = (HookManager)RuntimeHelpers.GetUninitializedObject(typeof(HookManager));
            Set(manager, "_explorerWatcher", fixture.Watcher);
            Set(manager, "_sessionShortcuts", shortcuts);
            await body(fixture, manager, shortcuts.Dispatch);
        });

    private sealed class ShortcutConfiguration : IExplorerSessionShortcutHook
    {
        public ExplorerShortcutDispatch Dispatch { get; } = new();
        public uint SettingsUiProcessId { get; set; }
        public event Action<string>? Failed { add { } remove { } }
        public void Configure(ExplorerShortcut? group, ExplorerShortcut? tab)
        {
            Dispatch.Group = group;
            Dispatch.Tab = tab;
            if (group == null && tab == null) Dispatch.Reset();
        }
        public void Dispose() { }
    }

    private static T Get<T>(object owner, string name) =>
        (T)owner.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(owner)!;

    private static void Set(object owner, string name, object value) =>
        owner.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(owner, value);

    private sealed class FakeHook : IHook
    {
        public bool FailStart { get; init; }
        public bool FailStop { get; init; }
        public int Starts { get; private set; }
        public bool IsHookActive { get; set; }

        public void StartHook()
        {
            if (FailStart) throw new InvalidOperationException("The input observer could not be installed.");
            Starts++;
            IsHookActive = true;
        }

        public void StopHook()
        {
            if (FailStop) throw new InvalidOperationException("The input observer could not be removed.");
            IsHookActive = false;
        }

        public void Dispose() { }
    }
}
