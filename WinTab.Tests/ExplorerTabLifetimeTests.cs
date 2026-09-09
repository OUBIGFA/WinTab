using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Interop;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.Interop;
using WinTab.Models;
using WinTab.WinAPI;

internal static class ExplorerTabLifetimeTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("reopened first tabs replace retired handle owners repeatedly", ReopenedFirstTabsReplaceRetiredOwners);
        yield return ("a live first tab cannot be stolen by another registration", LiveTabOwnerIsPreserved);
        yield return ("duplicate first-tab registration finishes without waiting for a new handle", DuplicateFirstTabRegistrationDoesNotWait);
        yield return ("three-tab registrations do not accumulate duplicate records", () => RepeatedTabRegistrationsStayBounded(3));
        yield return ("five-tab registrations do not accumulate duplicate records", () => RepeatedTabRegistrationsStayBounded(5));
        yield return ("a duplicate first-tab record cannot hide the reusable original", DuplicateRecordDoesNotHideOriginal);
        yield return ("tab lookup reads the shell catalog on its owning thread", TabLookupUsesShellThread);
        yield return ("a tab from another window cannot be published", ForeignTabCannotBePublished);
        yield return ("retired windows cannot satisfy a reuse search", RetiredWindowCannotBeReused);
        yield return ("unavailable Explorer reads do not prevent first-tab reuse", UnavailableReadDoesNotPreventReuse);
        yield return ("failed event detachment still removes a closed window", FailedEventDetachmentStillRemovesWindow);
        yield return ("navigation tracking starts before waiting for a tab handle", EventsAreHookedBeforeWaiting);
        yield return ("a late first tab is discovered before deciding to create another", LateFirstTabIsRecovered);
        yield return ("recycled tab handles do not inherit an earlier tab's owner", RecycledTabReplacesRetiredOwner);
        yield return ("retired tabs cannot match while their parent window remains open", RetiredTabCannotBeReused);
        yield return ("registration removes closed windows even without close events", RegistrationRemovesUnreportedClosedWindows);
        yield return ("registration revisits tracked tabs that were not ready", RegistrationRevisitsPendingTabs);
        yield return ("one unavailable Explorer window does not block other registrations", UnavailableWindowDoesNotBlockRegistration);
        yield return ("first-tab reuse survives repeated native window close and reopen", NativeWindowsCanCloseAndReopenRepeatedly);
        yield return ("a tab replaced during a location read cannot redirect reuse", ReplacedTabCannotRedirectReuse);
        yield return ("a first window published after the last event is discovered by polling", MissedFirstWindowIsDiscovered);
        yield return ("a new tab without a registration event is discovered by polling", MissedTabIsDiscovered);
        yield return ("empty-window intervals do not disable discovery of reopened first tabs", ReopenedWindowsAreDiscoveredAfterEmptyIntervals);
        yield return ("fully registered windows do not cause periodic COM scans", RegisteredWindowsDoNotRescanCatalog);
        yield return ("an empty desktop does not cause periodic COM scans", EmptyDesktopDoesNotRescanCatalog);
        yield return ("a window without Explorer tabs does not cause periodic COM scans", WindowWithoutTabsDoesNotRescanCatalog);
        yield return ("a disconnected catalog is retired without an Explorer process exit", () => DisconnectedCatalogIsRetired(new COMException("Disconnected", unchecked((int)0x80010108))));
        yield return ("an unavailable catalog server is retired without an Explorer process exit", () => DisconnectedCatalogIsRetired(new COMException("Server unavailable", unchecked((int)0x800706BA))));
        yield return ("a released catalog wrapper is retired without an Explorer process exit", () => DisconnectedCatalogIsRetired(new InvalidComObjectException("Released catalog")));
        yield return ("a temporarily busy catalog does not discard its live connection", BusyCatalogRemainsConnected);
        yield return ("a disconnected tab lookup retires the connection without a process exit", DisconnectedTabLookupRetiresConnection);
        yield return ("a disconnected window registration reaches connection recovery", DisconnectedWindowRegistrationRetiresConnection);
        yield return ("a completed merge can close its source after the work deadline", () => CompletedMergeFinishesSafely(false, 1));
        yield return ("completed merge cleanup still respects user cancellation", () => CompletedMergeFinishesSafely(true, 1));
        yield return ("completed merge cleanup leaves a source with extra tabs open", () => CompletedMergeFinishesSafely(false, 2));
        yield return ("retrying an existing window registration never conceals it", RetryingRegistrationLeavesExistingWindowVisible);
        yield return ("a closing tab with an unavailable selection still unregisters normally", ClosingTabDoesNotRequireItsDisconnectedSelection);
        yield return ("a live selection disconnect triggers recovery before the next folder open", LiveSelectionDisconnectRetiresConnection);
    }

    private static Task ReopenedFirstTabsReplaceRetiredOwners() => WithFixture(fixture =>
    {
        for (var cycle = 0; cycle < 250; cycle++)
        {
            var browser = fixture.AddBrowser(out var info);
            Check.That(fixture.Publish(browser, fixture.Window.FirstTab),
                "A reopened first tab must replace an expired owner of the same native handle.");
            Check.Equal(1, fixture.Count, "Closed-window records must not accumulate across reopen cycles.");
            Check.That(ReferenceEquals(browser, fixture.Search()), "Reuse must resolve the current first tab.");
            info.Identity.Release();
        }

        return Task.CompletedTask;
    });

    private static Task LiveTabOwnerIsPreserved() => WithFixture(fixture =>
    {
        var original = fixture.AddBrowser(out var originalInfo, fixture.Window.FirstTab);
        fixture.Invoke("HookWindowEvents", original, originalInfo);
        var duplicate = fixture.AddBrowser(out var duplicateInfo, failDetach: true);
        fixture.Invoke("HookWindowEvents", duplicate, duplicateInfo);
        Check.That(fixture.Publish(duplicate, fixture.Window.FirstTab), "A repeated report of a live tab must finish instead of retrying.");
        Check.That(ReferenceEquals(original, fixture.Search()), "The original live tab must remain reusable.");
        Check.That(originalInfo.EventsHooked && !originalInfo.Closed, "Removing a duplicate must retain the original navigation tracking.");
        Check.That(duplicateInfo.Closed && !duplicateInfo.EventsHooked, "Duplicate event subscriptions must be retired even if detaching fails.");
        Check.Equal(1, fixture.Count, "Only one record may remain for the same live tab.");
        return Task.CompletedTask;
    });

    private static Task DuplicateFirstTabRegistrationDoesNotWait() => WithFixture(async fixture =>
    {
        var original = fixture.AddBrowser(out var originalInfo, fixture.Window.FirstTab);
        fixture.Invoke("HookWindowEvents", original, originalInfo);
        var duplicate = fixture.AddBrowser(out var duplicateInfo);
        var elapsed = Stopwatch.StartNew();

        await (Task)fixture.Invoke("RegisterIndependentWindowAsync", duplicate, duplicateInfo, fixture.Window.Handle)!;

        Check.Equal(1, fixture.Count, "An already registered tab must not remain in the retry queue.");
        Check.That(elapsed.ElapsedMilliseconds < 1_000, "A ready duplicate must not consume the two-second registration timeout.");
        Check.That(!(bool)fixture.Invoke("HasPendingTabRegistrations")!, "Duplicate reports must not keep periodic full scans running.");
        Check.That(ReferenceEquals(original, fixture.Search()), "Completing duplicate registration must retain the reusable original.");
    }, tabCount: 1);

    private static Task RepeatedTabRegistrationsStayBounded(int tabCount) => WithFixture(fixture =>
    {
        var tabs = ExplorerWindowDiscovery.GetAllExplorerTabs(fixture.Window.Handle).ToArray();
        var owners = new Dictionary<nint, object>();
        foreach (var tab in tabs)
            owners.Add(tab, fixture.AddBrowser(out _, tab));

        for (var round = 0; round < 100; round++)
        {
            foreach (var tab in tabs)
            {
                var duplicate = fixture.AddBrowser(out var duplicateInfo);
                Check.That(fixture.Publish(duplicate, tab), "A duplicate tab report must complete on its first attempt.");
                Check.Equal(tabCount, fixture.Count, "Repeated registration must not accumulate records as more folders are opened.");
                Check.That(duplicateInfo.Closed, "Duplicate records must no longer participate in registration or reuse.");
                Check.That(ReferenceEquals(owners[tab], fixture.Invoke("GetWindowByTabHandle", tab, fixture.Window.Handle)),
                    "Distinct native tabs must retain their own owners even when their folder paths match.");
            }
        }

        return Task.CompletedTask;
    }, tabCount);

    private static Task DuplicateRecordDoesNotHideOriginal() => WithFixture(fixture =>
    {
        fixture.AddBrowser(out _);
        var original = fixture.AddBrowser(out _, fixture.Window.FirstTab);

        Check.That(ReferenceEquals(original, fixture.Search()), "A duplicate encountered first must not hide the original cached match.");
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task TabLookupUsesShellThread() => WithFixture(async fixture =>
    {
        var ownerThread = Environment.CurrentManagedThreadId;
        var readThreads = new List<int>();
        fixture.OnCatalogRead = () => readThreads.Add(Environment.CurrentManagedThreadId);
        fixture.SetCatalog();

        await Task.Run(async () => await (Task)fixture.Invoke("FindShellWindowByTabHandle", fixture.Window.FirstTab, fixture.Window.Handle)!);

        Check.That(readThreads.Count > 0, "An uncached tab lookup must query the shell catalog.");
        Check.That(readThreads.All(thread => thread == ownerThread),
            "Polling from a background thread must not create a second shell wrapper outside the registration thread.");
    });

    private static Task ForeignTabCannotBePublished() => WithFixture(fixture =>
    {
        using var other = new ExplorerTabActivationTests.ActivationWindow();
        var browser = fixture.AddBrowser(out _);
        Check.That(!fixture.Publish(browser, other.FirstTab), "A browser must not claim another window's tab.");
        return Task.CompletedTask;
    });

    private static Task RetiredWindowCannotBeReused() => WithFixture(fixture =>
    {
        fixture.AddBrowser(out var info, fixture.Window.FirstTab);
        info.Identity.Release();
        Check.That(fixture.Search() == null, "A closed window must not match even if its cached folder still matches.");
        return Task.CompletedTask;
    });

    private static Task UnavailableReadDoesNotPreventReuse() => WithFixture(async fixture =>
    {
        var browser = fixture.AddBrowser(out _, fixture.Window.FirstTab, unavailableLocation: true);
        Check.That(ReferenceEquals(browser, fixture.Search()), "A known first tab must survive a temporary COM read failure.");
        fixture.Window.SetActive(1);
        var reused = await (Task<bool>)fixture.Invoke("OpenTabNavigateWithSelection",
            new WindowRecord(Fixture.Location), fixture.Window.Handle)!;
        Check.That(reused, "The full reuse flow must succeed without re-reading a known folder.");
        Check.Equal(fixture.Window.FirstTab, fixture.Window.ActiveTab, "Reuse must activate the original first tab.");
        Check.Equal(2, ExplorerWindowDiscovery.GetAllExplorerTabs(fixture.Window.Handle).Count(),
            "First-tab reuse must not create an extra tab.");
    });

    private static Task FailedEventDetachmentStillRemovesWindow() => WithFixture(fixture =>
    {
        var browser = fixture.AddBrowser(out var info, fixture.Window.FirstTab, failDetach: true);
        info.EventsHooked = true;
        var handler = typeof(WindowInfo).GetProperty("OnQuitHandler")!;
        handler.SetValue(info, Delegate.CreateDelegate(handler.PropertyType,
            typeof(ExplorerTabLifetimeTests).GetMethod(nameof(OnQuit), BindingFlags.Static | BindingFlags.NonPublic)!));

        fixture.Invoke("RemoveWindowAndUnhookEvents", browser, info, true, true);

        Check.Equal(0, fixture.Count, "A failed COM unsubscribe must not retain the closed window or its tab handle.");
        Check.That(info.Closed && !info.EventsHooked, "The closed window must stop participating in event tracking.");
        return Task.CompletedTask;
    });

    private static Task EventsAreHookedBeforeWaiting() => WithFixture(async fixture =>
    {
        var browser = fixture.AddBrowser(out var info);
        var registration = (Task)fixture.Invoke("RegisterIndependentWindowAsync", browser, info, fixture.Window.Handle)!;
        try
        {
            Check.That(info.EventsHooked, "Navigation and close events must be tracked while the first tab is still becoming ready.");
        }
        finally
        {
            fixture.Publish(browser, fixture.Window.FirstTab);
            await registration;
        }
    });

    private static Task LateFirstTabIsRecovered() => WithFixture(async fixture =>
    {
        var browser = fixture.AddBrowser(out _);
        Check.That(ReferenceEquals(browser, fixture.Search()),
            "A first tab that became ready after registration must be discovered before creating a duplicate.");
        var reused = await (Task<bool>)fixture.Invoke("OpenTabNavigateWithSelection",
            new WindowRecord(Fixture.Location), fixture.Window.Handle)!;
        Check.That(reused, "The late first tab must participate in the normal reuse flow.");
        Check.Equal(1, ExplorerWindowDiscovery.GetAllExplorerTabs(fixture.Window.Handle).Count(),
            "Recovering a late registration must not add another tab.");
    }, tabCount: 1);

    private static Task RecycledTabReplacesRetiredOwner() => WithFixture(fixture =>
    {
        fixture.AddBrowser(out _, fixture.Window.FirstTab);
        for (var cycle = 0; cycle < 250; cycle++)
        {
            WindowIdentity.Capture(fixture.Window.FirstTab).Release();
            var replacement = fixture.AddBrowser(out _);
            Check.That(fixture.Publish(replacement, fixture.Window.FirstTab),
                "A recycled child handle must belong to the new tab even when its parent window remains open.");
            Check.Equal(1, fixture.Count, "Retired tabs must not accumulate while their parent stays open.");
            Check.That(ReferenceEquals(replacement, fixture.Search()), "Reuse must select the current tab owner.");
        }
        return Task.CompletedTask;
    });

    private static Task RetiredTabCannotBeReused() => WithFixture(fixture =>
    {
        fixture.AddBrowser(out _, fixture.Window.FirstTab);
        WindowIdentity.Capture(fixture.Window.FirstTab).Release();
        Check.That(fixture.Search() == null, "A retired child tab must not match just because its parent remains open.");
        return Task.CompletedTask;
    });

    private static void OnQuit() { }

    private static Task ReplacedTabCannotRedirectReuse() => WithFixture(fixture =>
    {
        var locationRead = false;
        fixture.AddBrowser(out var originalInfo, fixture.Window.FirstTab, readLocation: () =>
        {
            locationRead = true;
            WindowIdentity.Capture(fixture.Window.FirstTab).Release();
            var replacement = fixture.AddBrowser(out var replacementInfo);
            replacementInfo.Location = @"C:\Different-folder";
            Check.That(fixture.Publish(replacement, fixture.Window.FirstTab), "The fixture must replace the tab during the live read.");
        });
        originalInfo.Location = @"C:\Previous-folder";
        Check.That(fixture.Search() == null, "A live read from a retired tab must not redirect reuse into a different replacement tab.");
        Check.That(locationRead, "The replacement must actually occur during the live location read.");
        return Task.CompletedTask;
    });

    private static Task NativeWindowsCanCloseAndReopenRepeatedly() => WithFixture(fixture =>
    {
        var elapsed = Stopwatch.StartNew();
        for (var cycle = 0; cycle < 100; cycle++)
        {
            var browser = fixture.AddBrowser(out _);
            fixture.SetCatalog(browser);
            fixture.Invoke("AdoptNewShellWindows");
            Check.That(fixture.Publish(browser, fixture.Window.FirstTab), "Each reopened native window must publish its first tab.");
            Check.That(ReferenceEquals(browser, fixture.Search()), "Each reopened native window must reuse its own first tab.");
            Check.Equal(1, fixture.Count, "Native window teardown must not leave old registrations behind.");
            fixture.ReopenWindow();
        }
        fixture.SetCatalog();
        fixture.Invoke("AdoptNewShellWindows");
        Check.Equal(0, fixture.Count, "Closing the last native window must leave no stale registration.");
        Console.WriteLine($"Native window reopen endurance: 100 cycles in {elapsed.ElapsedMilliseconds} ms");
        return Task.CompletedTask;
    });

    private static Task RegistrationRemovesUnreportedClosedWindows() => WithFixture(fixture =>
    {
        fixture.AddBrowser(out var info, fixture.Window.FirstTab);
        info.Identity.Release();
        fixture.SetCatalog();
        fixture.Invoke("AdoptNewShellWindows");
        Check.Equal(0, fixture.Count, "Registration must clear a closed window even when Explorer omitted its close event.");
        return Task.CompletedTask;
    });

    private static Task RegistrationRevisitsPendingTabs() => WithFixture(fixture =>
    {
        var browser = fixture.AddBrowser(out _);
        fixture.SetCatalog(browser);
        var registrations = (System.Collections.ICollection)fixture.Invoke("AdoptNewShellWindows")!;
        Check.Equal(1, registrations.Count, "An incomplete first-tab registration must not be skipped forever.");
        return Task.CompletedTask;
    });

    private static Task UnavailableWindowDoesNotBlockRegistration() => WithFixture(fixture =>
    {
        var available = fixture.AddBrowser(out _, track: false);
        var unavailable = fixture.AddBrowser(out _, unreadableHandle: true, track: false);
        fixture.SetCatalog(available, unavailable);
        var registrations = (System.Collections.ICollection)fixture.Invoke("AdoptNewShellWindows")!;
        Check.Equal(1, registrations.Count, "A temporary COM failure in one window must not discard the other ready window.");
        Check.Equal(1, fixture.Count, "Only the available window should be adopted.");
        return Task.CompletedTask;
    });

    private static Task MissedFirstWindowIsDiscovered() => WithFixture(async fixture =>
    {
        fixture.SetCatalog();
        await (Task)fixture.Invoke("ProcessRegisteredShellWindowsAsync")!;
        var browser = fixture.AddBrowser(out _, track: false);
        fixture.SetCatalog(browser);

        Check.That(await fixture.PollForRegistrationAsync(fixture.Window.Handle),
            "A visible first window must be revisited even when its catalog entry appeared after the final event.");
        fixture.Invoke("AdoptNewShellWindows");
        Check.That(fixture.Publish(browser, fixture.Window.FirstTab), "The late first tab must become registered.");
        Check.That(ReferenceEquals(browser, fixture.Search()), "The late first tab must become reusable.");
    }, tabCount: 1);

    private static Task MissedTabIsDiscovered() => WithFixture(async fixture =>
    {
        fixture.AddBrowser(out var firstInfo, fixture.Window.FirstTab);
        firstInfo.EventsHooked = true;

        Check.That(await fixture.PollForRegistrationAsync(fixture.Window.Handle),
            "A new native tab must trigger registration even while every previously known tab remains valid.");
    });

    private static Task ReopenedWindowsAreDiscoveredAfterEmptyIntervals() => WithFixture(async fixture =>
    {
        for (var cycle = 0; cycle < 25; cycle++)
        {
            fixture.AddBrowser(out var previous, fixture.Window.FirstTab);
            previous.EventsHooked = true;
            fixture.ReopenWindow();
            fixture.SetCatalog();
            fixture.Invoke("AdoptNewShellWindows");
            Check.Equal(0, fixture.Count, "The old window must be fully retired before testing rediscovery.");
            Check.That(await fixture.PollForRegistrationAsync(fixture.Window.Handle),
                "An empty registry must not make a reopened first window invisible to periodic discovery.");
        }
    }, tabCount: 1);

    private static Task RegisteredWindowsDoNotRescanCatalog() => WithFixture(async fixture =>
    {
        foreach (var tab in ExplorerWindowDiscovery.GetAllExplorerTabs(fixture.Window.Handle))
        {
            fixture.AddBrowser(out var info, tab);
            info.EventsHooked = true;
        }
        fixture.SetCatalog();
        fixture.OnCatalogRead = () => throw new InvalidOperationException("Healthy polling must not read the COM catalog.");
        for (var iteration = 0; iteration < 100; iteration++)
        {
            Check.That(!await fixture.PollForRegistrationAsync(fixture.Window.Handle),
                "Healthy windows must not enqueue redundant full catalog scans.");
        }
    });

    private static Task EmptyDesktopDoesNotRescanCatalog() => WithFixture(async fixture =>
    {
        fixture.SetCatalog();
        fixture.OnCatalogRead = () => throw new InvalidOperationException("An empty desktop must not read the COM catalog.");
        Check.That(!await fixture.PollForRegistrationAsync(), "No windows means no catalog scan is needed.");
    });

    private static Task DisconnectedCatalogIsRetired(Exception failure) => WithFixture(async fixture =>
    {
        fixture.SetCatalog();
        fixture.MarkShellConnected();
        fixture.OnCatalogRead = () => throw failure;

        await (Task)fixture.Invoke("ProcessRegisteredShellWindowsAsync")!;

        Check.That(!fixture.ShellConnected,
            "A dead catalog must be retired so the next process check reconnects even while explorer.exe stays alive.");
        Check.That(fixture.ShellCancellationRequested,
            "Work from the disconnected catalog must be cancelled before its connection is replaced.");
    });

    private static Task WindowWithoutTabsDoesNotRescanCatalog() => WithFixture(async fixture =>
    {
        using var window = new HwndSource(new HwndSourceParameters("WinTab empty window test")
        {
            WindowStyle = 0,
            PositionX = -32000,
            PositionY = -32000,
            Width = 1,
            Height = 1
        });
        Check.That(!await fixture.PollForRegistrationAsync(window.Handle),
            "An empty or unsupported Explorer frame must not keep full catalog scans running indefinitely.");
    });

    private static Task BusyCatalogRemainsConnected() => WithFixture(async fixture =>
    {
        fixture.SetCatalog();
        fixture.MarkShellConnected();
        fixture.OnCatalogRead = () => throw new COMException("Busy", unchecked((int)0x80010001));

        await (Task)fixture.Invoke("ProcessRegisteredShellWindowsAsync")!;

        Check.That(fixture.ShellConnected, "A rejected call must remain a retry, not a full connection reset.");
        Check.That(!fixture.ShellCancellationRequested, "A busy catalog must not cancel unrelated live tabs.");
        fixture.OnCatalogRead = null;
        await (Task)fixture.Invoke("ProcessRegisteredShellWindowsAsync")!;
        Check.That(fixture.ShellConnected, "The original connection must remain usable after the rejected call.");
    });

    private static Task DisconnectedTabLookupRetiresConnection() => WithFixture(async fixture =>
    {
        var browser = fixture.AddBrowser(out var info, fixture.Window.FirstTab,
            readLocation: () => throw new COMException("Disconnected tab", unchecked((int)0x80010108)));
        info.Location = null;
        fixture.SetCatalog(browser);
        fixture.MarkShellConnected();
        await (Task)fixture.Invoke("RunShellWorkAsync", (Func<Task>)(() =>
        {
            fixture.Search();
            return Task.CompletedTask;
        }))!;
        Check.That(!fixture.ShellConnected && fixture.ShellCancellationRequested,
            "A disconnected tab must retire stale objects rather than silently request another tab.");
    });

    private static Task DisconnectedWindowRegistrationRetiresConnection() => WithFixture(async fixture =>
    {
        var browser = fixture.AddBrowser(out var info, fixture.Window.FirstTab,
            readLocation: () => throw new COMException("Disconnected registration", unchecked((int)0x80010108)));
        info.Location = null;
        fixture.SetCatalog(browser);
        fixture.MarkShellConnected();
        await fixture.RegisterThroughWorkerAsync(browser, info);
        Check.That(!fixture.ShellConnected && fixture.ShellCancellationRequested,
            "A single-window registration must not hide a disconnect from the connection owner.");
    });

    private static Task CompletedMergeFinishesSafely(bool cancel, int tabCount) => WithFixture(async fixture =>
    {
        var browser = fixture.AddBrowser(out var info, fixture.Window.FirstTab);
        var source = info.Identity;
        var closed = await fixture.CloseCompletedMergeAfterDeadlineAsync(browser, info, cancel);
        if (!cancel && tabCount == 1)
        {
            Check.That(closed && !source.IsCurrent, "A completed target must not leave its source behind just because the earlier work deadline elapsed.");
        }
        else
        {
            Check.That(!closed && source.IsCurrent, "Completion must never bypass cancellation or close a source containing other tabs.");
        }
    }, tabCount);

    private static Task RetryingRegistrationLeavesExistingWindowVisible() => WithFixture(async fixture =>
    {
        var reads = 0;
        var wasHidden = false;
        var browser = fixture.AddBrowser(out var info, fixture.Window.FirstTab, readLocation: () =>
        {
            reads++;
            wasHidden |= ExplorerWindowVisibility.Contains(fixture.Window.Handle) ||
                WinApi.GetLayeredWindowAttributes(fixture.Window.Handle, out _, out var alpha, out _) && alpha == 0;
        });
        info.Location = null;
        fixture.SetCatalog(browser);
        fixture.MarkShellConnected();
        fixture.EnableMerging();
        await (Task)fixture.Invoke("ProcessRegisteredShellWindowsAsync")!;
        Check.That(reads > 0 && info.EventsHooked, "The existing window must actually complete its pending registration.");
        Check.That(!wasHidden, "Retrying a known window must never hide it or send it through the new-window merge flow.");
        Check.Equal(1, fixture.Count, "Registration retries must preserve the original window record.");
    }, tabCount: 1);

    private static Task ClosingTabDoesNotRequireItsDisconnectedSelection() => WithFixture(fixture =>
    {
        var browser = fixture.AddBrowser(out var info, fixture.Window.FirstTab,
            documentFailure: new COMException("The closing view has disconnected", unchecked((int)0x80010108)));
        fixture.Invoke("HookWindowEvents", browser, info);
        var closeHandler = (Delegate?)typeof(WindowInfo).GetProperty("OnQuitHandler")!.GetValue(info);
        Check.That(closeHandler != null, "The test must deliver the real close event handler.");
        try { closeHandler!.DynamicInvoke(); }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
        }
        Check.Equal(0, fixture.Count, "A closing tab must unregister even when its optional selection is no longer readable.");
        return Task.CompletedTask;
    });

    private static Task LiveSelectionDisconnectRetiresConnection() => WithFixture(async fixture =>
    {
        var browser = fixture.AddBrowser(out _, fixture.Window.FirstTab,
            documentFailure: new COMException("The active view disconnected", unchecked((int)0x80010108)));
        fixture.SetCatalog(browser);
        fixture.MarkShellConnected();
        await (Task)fixture.Invoke("RunShellWorkAsync", (Func<Task>)(() =>
        {
            fixture.Invoke("TryGetSelectedItems", browser);
            return Task.CompletedTask;
        }))!;
        Check.That(!fixture.ShellConnected && fixture.ShellCancellationRequested,
            "A live background read must trigger reconnection without waiting for another double-click.");
    });

    private static async Task WithFixture(Func<Fixture, Task> test, int tabCount = 2)
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(async () =>
        {
            using var fixture = new Fixture(scheduler, tabCount);
            await test(fixture);
        }, CancellationToken.None, TaskCreationOptions.None, scheduler).Unwrap();
    }

    private sealed class Fixture : IDisposable
    {
        public const string Location = @"C:\WinTab-lifetime";
        private readonly CancellationTokenSource _lifetime = new();
        private readonly SemaphoreSlim _openLock = new(1);
        private readonly TabStripHitTester _tabStrip = new();
        private readonly MergeSourceConcealPulse _concealPulse = new();
        private readonly ShellPathComparer _pathComparer = new();
        private readonly Timer _mergeTimer = new(_ => { });
        private readonly ExplorerWatcher _watcher;
        private readonly object _dictionary;
        private readonly Type _dictionaryType;
        public ExplorerTabActivationTests.ActivationWindow Window { get; private set; }
        public int Count => (int)_dictionaryType.GetProperty("Count")!.GetValue(_dictionary)!;
        public Action? OnCatalogRead { get; set; }
        public bool ShellConnected => (int)typeof(ExplorerWatcher).GetField("_mainExplorerProcessId", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(_watcher)! != 0;
        public bool ShellCancellationRequested => _lifetime.IsCancellationRequested;

        public Fixture(StaTaskScheduler scheduler, int tabCount)
        {
            Window = new ExplorerTabActivationTests.ActivationWindow(tabCount);
            _watcher = ExplorerTabActivationTests.CreateSelectionWatcher(_lifetime);
            _dictionaryType = typeof(ExplorerWatcher).GetField("_windowEntryDict", BindingFlags.Instance | BindingFlags.NonPublic)!.FieldType;
            _dictionary = Activator.CreateInstance(_dictionaryType)!;
            SetField("_windowEntryDict", _dictionary);
            SetField("_windowEntryDictLock", new object());
            SetField("_staTaskScheduler", scheduler);
            SetField("_shellPathComparer", _pathComparer);
            SetField("_hookLifetime", _lifetime);
            SetField("_mergeSafetyTimer", _mergeTimer);
            SetField("_locationResolver", new ExplorerLaunchLocationResolver());
            SetField("_getExplorerWindows", (Func<IEnumerable<nint>>)(() => []));
            SetField("_defaultLocation", "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}");
            SetField("_startupLocationCache", new ConcurrentDictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
            SetField("_closedWindows", new List<WindowRecord>());
            SetField("_closedWindowsLock", new object());
            SetField("_toOpenWindowsLock", _openLock);
            SetField("_reuseTabs", true);
            SetField("_processedHWnds", new ConcurrentDictionary<nint, WindowIdentity>());
            SetField("_hookedTopLevelUseCounts", new ConcurrentDictionary<nint, int>());
            SetField("_closingMergeSourceHWnds", new ConcurrentDictionary<nint, MergeOperation>());
            SetField("_mergeSourceHWnds", Activator.CreateInstance(typeof(ExplorerWatcher).GetField("_mergeSourceHWnds", BindingFlags.Instance | BindingFlags.NonPublic)!.FieldType)!);
            SetField("_mergeSourceConcealPulse", _concealPulse);
            SetField("<TabStrip>k__BackingField", _tabStrip);
        }

        public object AddBrowser(out WindowInfo info, nint? tab = null, bool unavailableLocation = false,
            bool failDetach = false, bool unreadableHandle = false, bool track = true, Action? readLocation = null,
            Exception? documentFailure = null)
        {
            var parentHandle = Window.Handle;
            var browser = ShellDispatchStub.Create(_dictionaryType.GetGenericArguments()[0], (method, arguments) =>
            {
                if (failDetach && method.StartsWith("remove_", StringComparison.Ordinal))
                    throw new COMException("The Explorer connection has already closed.");
                if (method == "get_LocationURL")
                    readLocation?.Invoke();
                return method switch
                {
                    "get_HWND" => unreadableHandle
                        ? throw new COMException("The Explorer window is no longer available.")
                        : (int)parentHandle,
                    "get_LocationURL" => unavailableLocation
                        ? throw new COMException("Explorer is temporarily unavailable.")
                        : "file:///C:/WinTab-lifetime",
                    "get_Document" => documentFailure != null ? throw documentFailure : null,
                    "get_Busy" => false,
                    "add_OnQuit" or "remove_OnQuit" or "add_NavigateComplete2" or "remove_NavigateComplete2" => null,
                    _ => throw new InvalidOperationException("Unexpected browser call: " + method)
                };
            });
            info = new WindowInfo
            {
                Identity = WindowIdentity.Capture(Window.Handle),
                HookedTopLevelHWnd = Window.Handle,
                Location = Location
            };
            if (track)
            {
                _dictionaryType.GetMethods().Single(method => method.Name == "Add" && method.GetParameters().Length == 3)
                    .Invoke(_dictionary, [browser, info, null]);
                if (tab.HasValue)
                    Check.That(Publish(browser, tab.Value), "The fixture's live tab must be registered before testing reuse.");
            }
            return browser;
        }

        public void SetCatalog(params object[] browsers)
        {
            var catalogType = typeof(ExplorerWatcher).GetField("_shellWindows", BindingFlags.Instance | BindingFlags.NonPublic)!.FieldType;
            SetField("_shellWindows", ShellDispatchStub.Create(catalogType, (method, arguments) =>
            {
                OnCatalogRead?.Invoke();
                return method switch
                {
                    "get_Count" => browsers.Length,
                    "Item" => browsers[Convert.ToInt32(arguments[0])],
                    _ => throw new InvalidOperationException("Unexpected catalog call: " + method)
                };
            }));
        }

        public void ReopenWindow()
        {
            Window.Dispose();
            Window = new ExplorerTabActivationTests.ActivationWindow();
        }

        public bool Publish(object browser, nint tab) => (bool)Invoke("TryPublishTabHandle", browser, tab)!;

        public async Task<bool> PollForRegistrationAsync(params nint[] windows)
        {
            var requested = false;
            var work = new CoalescingAsyncWork(() =>
            {
                requested = true;
                return Task.CompletedTask;
            });
            SetField("_getExplorerWindows", (Func<IEnumerable<nint>>)(() => windows));
            SetField("_registrationWork", work);
            SetField("_preExistingExplorerWindowsProtected", true);
            SetField("_mainExplorerProcessId", Environment.ProcessId);
            Invoke("CheckForMainExplorer", new object?[] { null });
            await work.WhenIdle;
            return requested;
        }

        public void MarkShellConnected()
        {
            SetField("_mainExplorerProcessId", Environment.ProcessId);
            SetField("_preExistingExplorerWindowsProtected", true);
        }

        public void EnableMerging()
        {
            SetField("_isForcingTabs", true);
            SetField("_preExistingExplorerWindowsProtected", true);
        }

        public Task RegisterThroughWorkerAsync(object browser, WindowInfo info)
        {
            SetField("_isForcingTabs", true);
            Invoke("PreventWindowHiding", Window.Handle);
            return (Task)Invoke("RunShellWorkAsync", (Func<Task>)(() =>
                (Task)Invoke("ProcessRegisteredShellWindowAsync", browser, info)!))!;
        }

        public async Task<bool> CloseCompletedMergeAfterDeadlineAsync(object browser, WindowInfo info, bool cancel)
        {
            SetField("_isForcingTabs", true);
            SetField("_preExistingExplorerWindowsProtected", true);
            Invoke("HideMergeSourceWindow", Window.Handle);
            using var work = new MergeOperation(info.Identity, 0, _lifetime.Token, () => true, 1);
            var context = (AsyncLocal<MergeOperation?>)typeof(ExplorerWatcher)
                .GetField("_currentMerge", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(_watcher)!;
            context.Value = work;
            try
            {
                try { await Task.Delay(Timeout.Infinite, work.Token).WaitAsync(TimeSpan.FromSeconds(2)); }
                catch (OperationCanceledException) { }
                if (cancel)
                    _lifetime.Cancel();
                try { return await (Task<bool>)Invoke("CloseMergedSourceWindowAsync", browser, Window.Handle)!; }
                catch (OperationCanceledException) { return false; }
            }
            finally
            {
                context.Value = null;
                Invoke("RecoverHiddenExplorerWindows", "test-completion-cleanup");
            }
        }

        public object? Search()
        {
            object?[] arguments = [Location, (nint)0, (nint)0, null];
            return (bool)Invoke("TrySearchForTab", arguments)! ? arguments[3] : null;
        }

        public object? Invoke(string name, params object?[] arguments)
        {
            try
            {
                return typeof(ExplorerWatcher).GetMethod(name, BindingFlags.Instance | BindingFlags.Static | BindingFlags.NonPublic)!
                    .Invoke(_watcher, arguments);
            }
            catch (TargetInvocationException exception) when (exception.InnerException != null)
            {
                ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
                throw;
            }
        }

        private void SetField(string name, object value) =>
            typeof(ExplorerWatcher).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(_watcher, value);

        public void Dispose()
        {
            _tabStrip.Dispose();
            _pathComparer.Dispose();
            _mergeTimer.Dispose();
            _concealPulse.Stop();
            Window.Dispose();
            _openLock.Dispose();
            _lifetime.Dispose();
        }
    }
}
