using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.Managers;
using WinTab.Models;
using WinTab.WinAPI;
using Fixture = ExplorerTabLifetimeTests.Fixture;

/// <summary>Exercises the watcher/native boundary with owned test windows, never real user Explorer windows.</summary>
internal static class ExplorerSessionNativeTests
{
    private const BindingFlags PrivateInstance = BindingFlags.Instance | BindingFlags.NonPublic;

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("session native restore remains active with merging and its lifetime disabled", IndependentFromMergeLifetime);
        yield return ("session native restore rejects a disabled restore toggle", () => GuardRejects((fixture, _) => Set(fixture.Watcher, "_restoreTabs", false)));
        yield return ("session native restore rejects a changed restore generation", () => GuardRejects((fixture, _) => Set(fixture.Watcher, "_sessionGeneration", 1)));
        yield return ("session native restore rejects a navigation in the initial tab", () => GuardRejects((_, info) => info.Location = @"C:\changed"));
        yield return ("session native restore rejects an extra user-created tab", () => GuardRejects((fixture, _) => fixture.Window.AddTab()));
        yield return ("session native restore rejects a recycled frame identity", () => GuardRejects((_, info) => info.Identity.Release()));
        yield return ("session native restore yields when another window takes foreground", ForegroundChangeCancels);
        yield return ("session native restore does not navigate or foreground an already active initial tab", InitialSelectionIsNonDestructive);
        yield return ("session native restore closes only the active placeholder and retains the restored tabs", () => CloseActivePlaceholder(0));
        yield return ("session native restore waits for a delayed placeholder close without resending it", () => CloseActivePlaceholder(300));
        yield return ("session native placeholder close leaves the shell callback thread responsive", CloseDoesNotBlockShellCallbacks);
        yield return ("session initial tab identity retries an unpublished strip on a background worker", InitialTabIdentityRetriesEmptyRead);
        yield return ("session native restore refuses to close a background placeholder", BackgroundPlaceholderIsNotClosed);
        yield return ("session watcher saves a closed single-tab frame through the store", ClosedSingletonIsStored);
        yield return ("session watcher keeps the saved group when a single-tab frame closes with single-tab restore off", ClosedSingletonKeepsGroupWhenOff);
        yield return ("session watcher uses actual close order rather than dictionary order", LastClosedFrameWins);
        yield return ("session watcher never captures a concealed merge source", MergeSourcesAreNotHistory);
        yield return ("session watcher skips restoration while another shown window exists", OtherWindowPreventsRestore);
        yield return ("session watcher excludes windows that existed before enabling restoration", ExistingWindowIsExcluded);
        yield return ("session watcher does not restore or capture a hidden preload", HiddenPreloadIsNotUserWindow);
        yield return ("session watcher registers an existing hidden preload when restoration is enabled", EnablingRestorationTracksExistingPreload);
        yield return ("session watcher cancels a restore when foreground leaves and returns during resolution", ForegroundRoundTripCancelsRestore);
        yield return ("session watcher does not let a window closed before its restore replace the saved group", PendingRestoreIsNotHistory);
        yield return ("session watcher treats the configured Explorer start folder as a normal launch", ConfiguredStartFolderIsNormalLaunch);
    }

    private static Task IndependentFromMergeLifetime() => WithNative(async (fixture, environment, _) =>
    {
        using var retiredMergeLifetime = new CancellationTokenSource();
        retiredMergeLifetime.Cancel();
        Set(fixture.Watcher, "_hookLifetime", retiredMergeLifetime);
        Set(fixture.Watcher, "_isForcingTabs", false);
        environment.EnsureUnchanged();
        Check.That(await environment.SelectTabAsync(environment.InitialTab),
            "Restoration must not depend on the merge toggle or the merge operation's cancelled lifetime.");
    });

    private static Task GuardRejects(Action<Fixture, WindowInfo> change) => WithNative((fixture, environment, info) =>
    {
        environment.EnsureUnchanged();
        change(fixture, info);
        Check.Throws<OperationCanceledException>(environment.EnsureUnchanged,
            "A stale or changed restore window must be rejected before another native command.");
        return Task.CompletedTask;
    });

    private static Task ForegroundChangeCancels() => WithNative((_, environment, _) =>
    {
        using var other = new ExplorerTabActivationTests.ActivationWindow(1);
        other.Show();
        Helper.RestoreWindowToForeground(other.Handle);
        Check.Equal(other.Handle, ExplorerNavigationAccess.ForegroundFrame(), "The owned second test frame must hold foreground.");
        Check.Throws<OperationCanceledException>(environment.EnsureUnchanged,
            "Restoration must not steal foreground back after the user leaves its window.");
        return Task.CompletedTask;
    });

    private static Task InitialSelectionIsNonDestructive() => WithNative(async (fixture, environment, info) =>
    {
        info.SelectedItems = ["selected-file"];
        Check.That(await environment.SelectTabAsync(environment.InitialTab), "An already active initial tab needs no navigation.");
        Check.Equal(environment.InitialTab, fixture.Window.ActiveTab, "The native active handle must be unchanged.");
        Check.Equal("selected-file", info.SelectedItems[0], "The native initial selection must be retained.");
    });

    private static Task CloseActivePlaceholder(int delayMs) => WithRestoredNativeTabs(async (frame, environment) =>
    {
        frame.CloseTabDelayMs = delayMs;
        Check.That(await environment.CloseInitialTabAsync(), "The active placeholder close and its successor must both be confirmed.");
        environment.EnsureUnchanged();
        Check.That(!WindowIdentity.Read(environment.InitialTab).IsCurrent, "The placeholder must actually be gone.");
        Check.That(ExplorerWindowDiscovery.GetAllExplorerTabs(frame.Handle).ToHashSet().SetEquals([frame.TabAt(1), frame.TabAt(2)]),
            "Both restored native tabs must survive the placeholder close.");
        Check.That(!await environment.CloseInitialTabAsync(), "A completed close must never be sent again.");
        Check.Equal(1, frame.CloseTabCommandCount, "The native window must receive exactly one close command.");
    });

    private static Task CloseDoesNotBlockShellCallbacks() => WithRestoredNativeTabs(async (frame, environment) =>
    {
        var shellContext = SynchronizationContext.Current!;
        var placeholder = WindowIdentity.Capture(frame.Tab);
        var callback = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        environment.GetType().GetField("_closeSelectedTab", PrivateInstance)!.SetValue(environment, (Func<bool>)(() =>
        {
            // Explorer's close can wait for OnQuit to return on the ShellWindows owner's STA thread.
            shellContext.Post(_ => callback.TrySetResult(placeholder.IsCurrent &&
                WinApi.PostMessage(placeholder.Handle, WinApi.WM_COMMAND, 0xA021, 1)), null);
            var deadline = Environment.TickCount64 + 1_000;
            while (!callback.Task.IsCompleted && Environment.TickCount64 < deadline)
                Thread.Sleep(1); // Do not pump COM or Dispatcher messages inside the simulated UIA call.
            return callback.Task.IsCompletedSuccessfully && callback.Task.Result;
        }));
        Check.That(await environment.CloseInitialTabAsync(),
            "Closing the placeholder must leave the ShellWindows STA free to deliver Explorer's callback.");
        Check.Equal(1, frame.CloseTabCommandCount, "One successful UIA close must produce exactly one native close.");
        Check.That(!placeholder.IsCurrent, "Only the original placeholder must be removed.");
    });

    private static Task InitialTabIdentityRetriesEmptyRead() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        fixture.AddBrowser(out var info, fixture.Window.FirstTab);
        var cacheField = typeof(ExplorerWatcher).GetField("_sessionVisuals", PrivateInstance)!;
        var snapshotType = cacheField.FieldType.GetGenericArguments()[1];
        var ownerThread = Environment.CurrentManagedThreadId;
        var reads = 0;
        var wrongThread = 0;
        Func<WindowIdentity, object?> read = _ =>
        {
            if (Environment.CurrentManagedThreadId == ownerThread || Thread.CurrentThread.GetApartmentState() != ApartmentState.MTA)
                Interlocked.Exchange(ref wrongThread, 1);
            if (Interlocked.Increment(ref reads) == 1)
                return null; // Explorer has shown the frame but has not published its first tab item yet.
            return Activator.CreateInstance(snapshotType,
                [new[] { info.TabIdentity.Handle }, info.TabIdentity.Handle,
                    new[] { new SessionVisualTab("ready-tab-id", "initial", true) }, Environment.TickCount64]);
        };
        using var cache = (IDisposable)typeof(ExplorerSessionNativeTests).GetMethod(nameof(ReadingCache), BindingFlags.Static | BindingFlags.NonPublic)!
            .MakeGenericMethod(snapshotType).Invoke(null, [read])!;
        cacheField.SetValue(fixture.Watcher, cache);
        var id = await (Task<string>)fixture.Invoke("ReadInitialSessionTabIdAsync", info, CancellationToken.None)!;
        Check.Equal("ready-tab-id", id, "An initially empty UIA read must be retried once Explorer publishes the tab.");
        Check.That(reads >= 2, "The initially unpublished strip must be queried again, not cached as permanently absent.");
        Check.Equal(0, wrongThread, "Initial tab identity reads must use the bounded MTA workers, not the ShellWindows STA.");
    }, tabCount: 1);

    private static Task BackgroundPlaceholderIsNotClosed() => WithRestoredNativeTabs(async (frame, environment) =>
    {
        Check.That(await environment.SelectTabAsync(frame.TabAt(1)), "The owned restored tab must be active before attempting the unsafe close.");
        Check.That(!await environment.CloseInitialTabAsync(), "A background placeholder must never be sent Explorer's active-tab close command.");
        Check.Equal(0, frame.CloseTabCommandCount, "Refusing the close must send no native command.");
        Check.Equal(3, ExplorerWindowDiscovery.GetAllExplorerTabs(frame.Handle).Count(), "The placeholder and both restored tabs must remain.");
    });

    private static Task WithRestoredNativeTabs(Func<RemoteExplorerFrame, IExplorerSessionRestoreEnvironment, Task> test) =>
        ExplorerTabLifetimeTests.WithFixture(async fixture =>
        {
            using var frame = new RemoteExplorerFrame(visible: true, tabCount: 3);
            var browser = fixture.AddBrowser(out var info, frame.Tab, handle: frame.Handle);
            Set(fixture.Watcher, "_restoreTabs", true);
            fixture.SetExplorerWindows(frame.Handle);
            frame.SetActive(0);
            Helper.RestoreWindowToForeground(frame.Handle);
            var foreground = await Helper.DoUntilConditionAsync(ExplorerNavigationAccess.ForegroundFrame,
                active => active == frame.Handle, 1_000, 20);
            Check.Equal(frame.Handle, foreground, "Only the owned remote test frame may receive native close commands.");
            var type = typeof(ExplorerWatcher).GetNestedType("NativeSessionRestore", BindingFlags.NonPublic)!;
            var environment = (IExplorerSessionRestoreEnvironment)Activator.CreateInstance(type,
                [fixture.Watcher, browser, info, Fixture.Location, string.Empty, 0, CancellationToken.None])!;
            // Start at the native boundary after two successful appends; creation itself is covered separately.
            var expected = (List<WindowIdentity>)type.GetField("_expected", PrivateInstance)!.GetValue(environment)!;
            expected.AddRange([WindowIdentity.Capture(frame.TabAt(1)), WindowIdentity.Capture(frame.TabAt(2))]);
            var context = (AsyncLocal<MergeOperation?>)typeof(ExplorerWatcher).GetField("_currentMerge", PrivateInstance)!.GetValue(fixture.Watcher)!;
            context.Value = (MergeOperation)type.GetProperty("Operation")!.GetValue(environment)!;
            type.GetField("_closeSelectedTab", PrivateInstance)!.SetValue(environment,
                (Func<bool>)(() => WinApi.PostMessage(frame.Tab, WinApi.WM_COMMAND, 0xA021, 1)));
            try { await test(frame, environment); }
            finally
            {
                context.Value = null;
                ((IDisposable)environment).Dispose();
                Set(fixture.Watcher, "_restoreTabs", false);
            }
        }, tabCount: 1);

    private static Task WithNative(Func<Fixture, IExplorerSessionRestoreEnvironment, WindowInfo, Task> test) =>
        ExplorerTabLifetimeTests.WithFixture(async fixture =>
        {
            var browser = fixture.AddBrowser(out var info, fixture.Window.FirstTab);
            Set(fixture.Watcher, "_restoreTabs", true);
            fixture.Window.Show();
            Helper.RestoreWindowToForeground(fixture.Window.Handle);
            Check.Equal(fixture.Window.Handle, ExplorerNavigationAccess.ForegroundFrame(), "Only the owned test window may receive native session commands.");
            var type = typeof(ExplorerWatcher).GetNestedType("NativeSessionRestore", BindingFlags.NonPublic)!;
            var environment = (IExplorerSessionRestoreEnvironment)Activator.CreateInstance(type,
                [fixture.Watcher, browser, info, Fixture.Location, string.Empty, 0, CancellationToken.None])!;
            var context = (AsyncLocal<MergeOperation?>)typeof(ExplorerWatcher).GetField("_currentMerge", PrivateInstance)!.GetValue(fixture.Watcher)!;
            context.Value = (MergeOperation)type.GetProperty("Operation")!.GetValue(environment)!;
            try { await test(fixture, environment, info); }
            finally
            {
                context.Value = null;
                ((IDisposable)environment).Dispose();
                Set(fixture.Watcher, "_restoreTabs", false);
            }
        }, tabCount: 1);

    private static Task ClosedSingletonIsStored() => WithCapture(async (fixture, store) =>
    {
        using var frame = new RemoteExplorerFrame(visible: true, explorerClass: true);
        var handle = frame.Handle;
        fixture.AddBrowser(out var info, frame.Tab, handle: handle);
        info.Location = @"C:\single-session";
        fixture.SetExplorerWindows(handle);
        fixture.Invoke("CaptureExplorerSessions");
        Check.That(store.Snapshot == null, "An open window is not the last closed session yet.");
        frame.Dispose();
        fixture.SetExplorerWindows();
        fixture.Invoke("NotifySessionWindowDestroyed", handle);
        fixture.Invoke("CompleteClosedSessions");
        Check.Equal(@"C:\single-session", store.Snapshot!.Locations[0], "Closing a singleton must immediately publish its complete session in memory.");
        Check.That(await store.FlushAsync(), "The singleton must also persist successfully.");
    });

    private static Task ClosedSingletonKeepsGroupWhenOff() => WithCapture(async (fixture, store) =>
    {
        Check.That(await store.SaveAsync(new ExplorerSession
        {
            Locations = [@"C:\group-a", @"C:\group-b"], ActiveTabIndex = 0, OrderVerified = true
        }), "The saved group must be stored first.");
        using var frame = new RemoteExplorerFrame(visible: true, explorerClass: true);
        var handle = frame.Handle;
        fixture.AddBrowser(out var info, frame.Tab, handle: handle);
        info.Location = @"C:\single-session";
        fixture.SetExplorerWindows(handle);
        fixture.Invoke("CaptureExplorerSessions");
        frame.Dispose();
        fixture.SetExplorerWindows();
        fixture.Invoke("NotifySessionWindowDestroyed", handle);
        fixture.Invoke("CompleteClosedSessions");
        Check.That(store.Snapshot!.Locations.SequenceEqual([@"C:\group-a", @"C:\group-b"]),
            "Closing a single-tab window must not replace the saved group while single-tab restore is off.");
    }, singleTab: false);

    private static Task LastClosedFrameWins() => WithCapture(async (fixture, store) =>
    {
        using var first = new RemoteExplorerFrame(visible: true, explorerClass: true);
        using var second = new RemoteExplorerFrame(visible: true, explorerClass: true);
        var firstHandle = first.Handle;
        var secondHandle = second.Handle;
        fixture.AddBrowser(out var firstInfo, first.Tab, handle: firstHandle);
        fixture.AddBrowser(out var secondInfo, second.Tab, handle: secondHandle);
        firstInfo.Location = @"C:\closed-last";
        secondInfo.Location = @"C:\closed-first";
        fixture.SetExplorerWindows(firstHandle, secondHandle);
        fixture.Invoke("CaptureExplorerSessions");
        second.Dispose();
        fixture.Invoke("NotifySessionWindowDestroyed", secondHandle);
        first.Dispose();
        fixture.Invoke("NotifySessionWindowDestroyed", firstHandle);
        fixture.SetExplorerWindows();
        fixture.Invoke("CompleteClosedSessions");
        Check.Equal(@"C:\closed-last", store.Snapshot!.Locations[0], "WinEvent close ordering, not frame registration/dictionary ordering, must select the last group.");
        Check.Equal(1, store.Snapshot.Locations.Length, "Independent windows must never be merged into one historical group.");
        Check.That(await store.FlushAsync(), "Only the last closed group must persist.");
    });

    private static Task MergeSourcesAreNotHistory() => WithCapture((fixture, store) =>
    {
        using var source = new RemoteExplorerFrame(visible: true, explorerClass: true);
        fixture.AddBrowser(out _, source.Tab, handle: source.Handle);
        fixture.SetExplorerWindows(source.Handle);
        fixture.EnableMerging();
        fixture.Invoke("HideMergeSourceWindow", source.Handle);
        fixture.Invoke("CaptureExplorerSessions");
        source.Dispose();
        fixture.SetExplorerWindows();
        fixture.Invoke("CompleteClosedSessions");
        Check.That(store.Snapshot == null, "Closing a window consumed by automatic merging must not replace session history.");
        return Task.CompletedTask;
    });

    private static Task OtherWindowPreventsRestore() => WithCapture(async (fixture, store) =>
    {
        using var first = new RemoteExplorerFrame(visible: true, explorerClass: true);
        using var second = new RemoteExplorerFrame(visible: true, explorerClass: true);
        var browser = fixture.AddBrowser(out var info, first.Tab, handle: first.Handle);
        fixture.SetExplorerWindows(first.Handle, second.Handle);
        Check.That(!await (Task<bool>)fixture.Invoke("TryRestoreNewExplorerWindowAsync", browser, info, false)!,
            "A second shown Explorer frame must disqualify restoration before any delayed work.");
        Check.That(store.Snapshot == null && first.IsAlive && second.IsAlive, "The exclusion must not close, mutate or save either window.");
    });

    private static Task ExistingWindowIsExcluded() => WithCapture(async (fixture, _) =>
    {
        using var frame = new RemoteExplorerFrame(visible: true, explorerClass: true);
        var browser = fixture.AddBrowser(out var info, frame.Tab, handle: frame.Handle);
        fixture.SetExplorerWindows(frame.Handle);
        fixture.Invoke("ExcludeExistingWindowsFromRestore");
        Check.That(!await (Task<bool>)fixture.Invoke("TryRestoreNewExplorerWindowAsync", browser, info, false)!,
            "Turning restoration on must not retroactively restore into an existing user window.");
    });

    private static Task HiddenPreloadIsNotUserWindow() => WithCapture(async (fixture, store) =>
    {
        using var frame = new RemoteExplorerFrame(visible: false, explorerClass: true);
        var browser = fixture.AddBrowser(out var info, frame.Tab, handle: frame.Handle);
        fixture.SetExplorerWindows(frame.Handle);
        fixture.Invoke("CaptureExplorerSessions");
        Check.That(!await (Task<bool>)fixture.Invoke("TryRestoreNewExplorerWindowAsync", browser, info, false)!,
            "A hidden preloaded frame must wait for an actual user-visible launch.");
        frame.Dispose();
        fixture.SetExplorerWindows();
        fixture.Invoke("CompleteClosedSessions");
        Check.That(store.Snapshot == null, "Discarding an unused hidden frame must never become session history.");
    });

    private static Task EnablingRestorationTracksExistingPreload() => WithCapture((fixture, _) =>
    {
        using var hidden = new RemoteExplorerFrame(visible: false, explorerClass: true);
        var browser = fixture.AddBrowser(out var info, hidden.Tab, handle: hidden.Handle);
        fixture.SetExplorerWindows(hidden.Handle);
        Set(fixture.Watcher, "_restoreTabs", false);
        Set(fixture.Watcher, "_sessionLifetime", new CancellationTokenSource());
        fixture.Watcher.SetRestoreTabs(true);
        var field = typeof(ExplorerWatcher).GetField("_preloadedSessionCandidates", PrivateInstance)!;
        var pending = (System.Collections.IEnumerable)field.GetValue(fixture.Watcher)!;
        Check.That(pending.Cast<object>().Any(),
            "An already-registered hidden preload must be checked on its next real show event.");
        Check.That(!ExplorerWindowDiscovery.IsShownExplorerWindow(hidden.Handle),
            "Merely enabling the feature must not expose or restore the hidden window.");
        return Task.CompletedTask;
    });

    private static Task ForegroundRoundTripCancelsRestore() => WithCapture(async (fixture, store) =>
    {
        Check.That(await store.SaveAsync(new ExplorerSession
        {
            Locations = [@"C:\saved-a", @"C:\saved-b"], ActiveTabIndex = 1, OrderVerified = true
        }), "The previous window must be stored before a restore is attempted.");
        using var frame = new RemoteExplorerFrame(visible: true, explorerClass: true);
        var observed = false;
        var browser = fixture.AddBrowser(out var info, frame.Tab, handle: frame.Handle, readLocation: () =>
        {
            if (observed) return;
            observed = true;
            fixture.Invoke("ObserveSessionForeground", frame.Handle);
            fixture.Invoke("ObserveSessionForeground", (nint)0);
            fixture.Invoke("ObserveSessionForeground", frame.Handle);
        });
        fixture.SetExplorerWindows(frame.Handle);
        Set(fixture.Watcher, "_restoreOnAnyFolder", true);
        await (Task<bool>)fixture.Invoke("TryRestoreNewExplorerWindowAsync", browser, info, false)!;
        Check.That(observed, "The launch location must have been read while the attempt was pending.");
        Check.Equal(0, frame.CloseTabCommandCount, "Returning to the window must not resume an interrupted restore.");
        Check.That(store.Snapshot!.Locations.SequenceEqual([@"C:\saved-a", @"C:\saved-b"]),
            "An interrupted attempt must not replace the saved window with a partial one.");
    });

    private static Task PendingRestoreIsNotHistory() => WithCapture(async (fixture, store) =>
    {
        Check.That(await store.SaveAsync(new ExplorerSession
        {
            Locations = [@"C:\saved-a", @"C:\saved-b"], ActiveTabIndex = 1, OrderVerified = true
        }), "The saved group must be stored before the new window opens.");
        using var frame = new RemoteExplorerFrame(visible: true, explorerClass: true);
        var handle = frame.Handle;
        var closed = false;
        // The user closes the new window while its launch location is still being resolved.
        var browser = fixture.AddBrowser(out var info, frame.Tab, handle: handle, readLocation: () =>
        {
            if (closed) return;
            closed = true;
            frame.Dispose();
        });
        fixture.SetExplorerWindows(handle);
        await (Task<bool>)fixture.Invoke("TryRestoreNewExplorerWindowAsync", browser, info, false)!;
        Check.That(closed, "The window must have been closed during the restore attempt.");
        fixture.SetExplorerWindows();
        fixture.Invoke("NotifySessionWindowDestroyed", handle);
        fixture.Invoke("CompleteClosedSessions");
        Check.That(store.Snapshot!.Locations.SequenceEqual([@"C:\saved-a", @"C:\saved-b"]),
            "A window that never finished its restore attempt must not become the last closed group.");
    });

    private static Task ConfiguredStartFolderIsNormalLaunch() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        Set(fixture.Watcher, "_defaultLocation", @"C:\Users\someone\Downloads");
        Check.That((bool)fixture.Invoke("IsSessionStartLocation", "file:///C:/Users/someone/Downloads")!,
            "A plain launch to the folder chosen in Explorer's options is a normal launch.");
        Check.That((bool)fixture.Invoke("IsSessionStartLocation", "shell:::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}")!,
            "Home remains a normal launch whatever the configured folder is.");
        Check.That(!(bool)fixture.Invoke("IsSessionStartLocation", @"C:\Users\someone\Documents")!,
            "Another folder is an explicit open.");
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task WithCapture(Func<Fixture, ExplorerSessionStore, Task> test, bool singleTab = true) =>
        ExplorerTabLifetimeTests.WithFixture(async fixture =>
        {
            var directory = Path.Combine(Path.GetTempPath(), "WinTab.Tests", Guid.NewGuid().ToString("N"));
            using var store = new ExplorerSessionStore(Path.Combine(directory, "session.json"));
            var work = new CoalescingAsyncWork(() => Task.CompletedTask);
            var cacheField = typeof(ExplorerWatcher).GetField("_sessionVisuals", PrivateInstance)!;
            using var cache = (IDisposable)typeof(ExplorerSessionNativeTests).GetMethod(nameof(EmptyCache), BindingFlags.Static | BindingFlags.NonPublic)!
                .MakeGenericMethod(cacheField.FieldType.GetGenericArguments()[1]).Invoke(null, null)!;
            foreach (var name in new[] { "_sessionTracker", "_sessionWindows", "_sessionClosedAt", "_sessionRestoreExclusions",
                "_restoringSessionWindows", "_preloadedSessionCandidates" })
            {
                var field = typeof(ExplorerWatcher).GetField(name, PrivateInstance)!;
                field.SetValue(fixture.Watcher, Activator.CreateInstance(field.FieldType));
            }
            Set(fixture.Watcher, "_sessionVisuals", cache);
            Set(fixture.Watcher, "_sessionStore", store);
            Set(fixture.Watcher, "_selectionWork", work);
            Set(fixture.Watcher, "_restoreSingleTab", singleTab);
            Set(fixture.Watcher, "_restoreTabs", true);
            try { await test(fixture, store); }
            finally
            {
                Set(fixture.Watcher, "_restoreTabs", false);
                await work.StopAsync();
                await store.FlushAsync();
                store.Dispose();
                TestCleanup.DeleteDirectory(directory);
            }
        }, tabCount: 1);

    private static BackgroundRefreshCache<WindowIdentity, T> ReadingCache<T>(Func<WindowIdentity, object?> read) where T : class =>
        new(identity => (T?)read(identity));

    private static BackgroundRefreshCache<WindowIdentity, T> EmptyCache<T>() where T : class => new(_ => null);

    private static void Set(ExplorerWatcher watcher, string name, object value) =>
        typeof(ExplorerWatcher).GetField(name, PrivateInstance)!.SetValue(watcher, value);
}
