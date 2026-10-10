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
        yield return ("session restore tolerates a late foreground callback after disposal", () => RetiredAttemptIgnoresCallbacks(false));
        yield return ("session restore tolerates a late navigation callback after disposal", () => RetiredAttemptIgnoresCallbacks(true));
        yield return ("session initial tab identity retries an unpublished strip on a background worker", InitialTabIdentityRetriesEmptyRead);
        yield return ("session watcher does not restore or capture a hidden preload", HiddenPreloadIsNotUserWindow);
        yield return ("session watcher registers an existing hidden preload when restoration is enabled", EnablingRestorationTracksExistingPreload);
        yield return ("session watcher treats the configured Explorer start folder as a normal launch", ConfiguredStartFolderIsNormalLaunch);
    }

    private static Task RetiredAttemptIgnoresCallbacks(bool navigation)
    {
        var type = typeof(ExplorerWatcher).GetNestedType("SessionRestoreAttempt", BindingFlags.NonPublic)!;
        using var attempt = (IDisposable)Activator.CreateInstance(type,
            [CancellationToken.None, true, default(WindowIdentity)])!;
        type.GetMethod("ArmNavigation")!.Invoke(attempt, null);
        var token = (CancellationToken)type.GetProperty("Token")!.GetValue(attempt)!;
        attempt.Dispose();
        // A concurrent dictionary enumerator can retain this attempt after its owner removes and disposes it.
        if (navigation)
            type.GetMethod("InitialTabNavigated")!.Invoke(attempt, null);
        else
            type.GetMethod("ObserveForeground")!.Invoke(attempt, [false]);
        Check.That(!token.IsCancellationRequested, "Retired callbacks must be inert rather than touch the disposed source.");
        return Task.CompletedTask;
    }

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
        fixture.Invoke("CompleteClosedSessions", false);
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

    internal static Task WithCapture(Func<Fixture, ExplorerSessionStore, Task> test, bool singleTab = true) =>
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
            Set(fixture.Watcher, "_captureSessions", true);
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
