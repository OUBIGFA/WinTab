using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.Models;
using Fixture = ExplorerTabLifetimeTests.Fixture;

/// <summary>
/// Replays the Control Panel-first launch reported in issue #6. All frames are message-only endpoints;
/// logical visibility is injected, and a failed target assertion stops before any foreground operation.
/// </summary>
internal static class ExplorerControlPanelTests
{
    private const string Category = "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}";
    private const string AllItems = "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}";
    private const string ThisPc = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
    private static readonly string[] ControlPanelLocations =
    [
        Category,
        Category + @"\0",
        Category + @"\8\::{7B81BE6A-CE2B-4676-A29E-EB907A5126C5}",
        AllItems,
        AllItems + @"\::{7007ACC7-3202-11D1-AAD2-00805FC1270E}",
        "shell:::{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}",
        "shell:ControlPanelFolder"
    ];

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        foreach (var location in ControlPanelLocations)
        {
            yield return ($"Control Panel isolation rejects merge and reopen targets: {location}", () => RejectsTargets(location));
            yield return ($"Control Panel isolation rejects cached tab reuse: {location}", () => RejectsReuse(location));
            yield return ($"Control Panel isolation rejects any-folder restore plans: {location}", () => RejectsRestorePlans(location));
        }
        yield return ("Control Panel isolation does not guess the purpose of an unregistered frame", UnregisteredFrameIsNotTarget);
        yield return ("Control Panel isolation keeps a normal folder available as a merge target", NormalFolderRemainsTarget);
        yield return ("Control Panel isolation preserves Control Panel when This PC opens afterwards", ControlPanelThenThisPc);
        yield return ("Control Panel isolation never conceals a known Control Panel window", KnownControlPanelIsNotConcealed);
        yield return ("Control Panel isolation does not consume an automatic restore attempt", DoesNotStartAutomaticRestore);
        yield return ("Control Panel isolation does not save Control Panel as a folder session", DoesNotCaptureControlPanel);
        yield return ("Control Panel isolation recognizes normalized namespace variants", RecognizesNamespaceVariants);
        yield return ("Control Panel isolation does not exclude ordinary folder names", PreservesOrdinaryLocations);
        yield return ("Control Panel isolation releases registered sources before automatic restore", SourcesBypassAutomaticRestore);
        yield return ("Control Panel isolation rejects direct and classic new-tab commands", RejectsTabCreation);
        yield return ("Control Panel isolation detects a live navigation before direct tab creation", DetectsLateNavigation);
        yield return ("Control Panel isolation rejects open requests before acquiring the merge lock", RejectsOpenRequest);
        yield return ("Control Panel isolation observes navigation changes without a stale exclusion", TracksNavigationChanges);
        yield return ("Control Panel isolation does not prevent a real first folder from restoring", DoesNotBlockFolderRestore);
        yield return ("Control Panel isolation never blocks WinEvent checks on the COM catalog lock", ChecksDoNotWaitForCatalog);
    }

    private static Task RejectsTargets(string location) => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var panel = new RemoteExplorerFrame(visible: false, explorerClass: true);
        AddPanel(fixture, panel, location);
        fixture.SetExplorerWindows(panel.Handle);
        fixture.SetShownExplorerWindows(panel.Handle);
        fixture.CacheMainWindow(panel.Handle);

        Check.That(!(bool)fixture.Invoke("IsStableMergeTargetWindow", panel.Handle, (nint)0)!,
            "A hooked Control Panel must not be considered a stable Explorer target.");
        Check.That(!(bool)fixture.Invoke("IsFallbackMergeTargetWindow", panel.Handle, (nint)0)!,
            "The fallback must not re-admit a Control Panel rejected by the preferred path.");
        Check.Equal((nint)0, (nint)fixture.Invoke("GetMainWindowHWnd", (nint)0, null)!,
            "A cached or enumerated Control Panel must not receive a newly opened folder.");
        Check.That(!(bool)fixture.Invoke("IsReopenTarget", panel.Handle)!,
            "Reopening a closed folder must not insert it into Control Panel either.");
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task RejectsReuse(string location) => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        var browser = fixture.AddBrowser(out var info, fixture.Window.FirstTab, locationUrl: () => location);
        info.Location = location;
        object?[] search = [location, (nint)0, (nint)0, null];
        Check.That(!(bool)fixture.Invoke("TrySearchForTab", search)!,
            "Control Panel must stay a native open, not a reusable folder tab.");
        Check.Equal((nint)0, (nint)search[2]!, "An excluded location must not publish a tab handle.");
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task RejectsRestorePlans(string location)
    {
        Check.That(!ExplorerSessionLocationPolicy.IsSupported(location),
            "Saved Control Panel addresses must remain unavailable to manual and automatic restoration.");
        var saved = new ExplorerSession { Locations = [@"C:\A", @"C:\B"] };
        foreach (var startup in new[] { false, true })
            Check.That(ExplorerSessionRestorePlan.Create(saved, location, startup, onAnyFolder: true,
                restoreSingleTab: true, new HashSet<int> { 0, 1 }) == null,
                "Even any-folder restoration must never append saved folders to Control Panel.");
        return Task.CompletedTask;
    }

    private static Task UnregisteredFrameIsNotTarget() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var unknown = new RemoteExplorerFrame(visible: false, explorerClass: true);
        fixture.SetExplorerWindows(unknown.Handle);
        fixture.SetShownExplorerWindows(unknown.Handle);
        Check.Equal((nint)0, (nint)fixture.Invoke("GetMainWindowHWnd", (nint)0, null)!,
            "CabinetWClass and ShellTabWindowClass alone cannot distinguish Control Panel from Explorer.");
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task NormalFolderRemainsTarget() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var folder = new RemoteExplorerFrame(visible: false, explorerClass: true);
        var browser = fixture.AddBrowser(out var info, folder.Tab, handle: folder.Handle);
        fixture.SetExplorerWindows(folder.Handle);
        fixture.SetShownExplorerWindows(folder.Handle);
        Check.Equal(folder.Handle, (nint)fixture.Invoke("GetMainWindowHWnd", (nint)0, null)!,
            "A registered folder with a known location remains available while its events are being attached.");
        fixture.Invoke("HookWindowEvents", browser, info);
        Check.That((bool)fixture.Invoke("IsStableMergeTargetWindow", folder.Handle, (nint)0)!,
            "Hooked folder windows must retain the stable-target fast path.");
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task ControlPanelThenThisPc() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var panel = new RemoteExplorerFrame(visible: false, explorerClass: true);
        using var folder = new RemoteExplorerFrame(visible: false, explorerClass: true);
        AddPanel(fixture, panel, AllItems);
        var browser = fixture.AddBrowser(out var info, folder.Tab, handle: folder.Handle, locationUrl: () => ThisPc);
        info.Location = ThisPc;
        fixture.SetExplorerWindows(folder.Handle, panel.Handle);
        fixture.SetShownExplorerWindows(folder.Handle, panel.Handle);
        fixture.CacheMainWindow(panel.Handle);
        fixture.EnableMerging();

        // Check the routing decision before invoking the owner, so a regression cannot send input or
        // attempt foreground activation even on these non-displayable endpoints.
        Check.Equal((nint)0, (nint)fixture.Invoke("GetMainWindowHWnd", folder.Handle, ThisPc)!,
            "Opening This PC with only Control Panel available must keep a separate folder window.");
        await (Task)fixture.Invoke("ProcessRegisteredShellWindowAsync", browser, info)!;
        Check.That(info.EventsHooked && panel.IsAlive && folder.IsAlive,
            "Both native windows must survive, and the new folder must remain registered for later opens.");
        Check.Equal(0, panel.CloseCommandCount + folder.CloseCommandCount, "Neither window may be closed by this open.");
        Check.Equal(0, panel.SwitchCommands.Count, "Control Panel must receive no tab-switch commands.");
        Check.Equal(0, fixture.MergeSourceCount, "The independent folder must not remain concealed.");
    }, tabCount: 1);

    private static Task KnownControlPanelIsNotConcealed() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var panel = new RemoteExplorerFrame(visible: false, explorerClass: true);
        var browser = fixture.AddBrowser(out var info, panel.Tab, handle: panel.Handle, locationUrl: () => AllItems);
        info.Location = AllItems;
        fixture.EnableMerging();
        try
        {
            fixture.Invoke("HideMergeSourceWindow", panel.Handle);
            Check.Equal(0, fixture.MergeSourceCount, "Known Control Panel windows must never enter concealment tracking.");
            Check.That(!(bool)fixture.Invoke("TryHideIncomingExplorerWindow", panel.Handle)!,
                "The early WinEvent path must also leave a known Control Panel alone.");
        }
        finally { fixture.Invoke("RecoverHiddenExplorerWindows", "control-panel-test-cleanup"); }
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task DoesNotStartAutomaticRestore() => ExplorerSessionNativeTests.WithCapture(async (fixture, _) =>
    {
        using var panel = new RemoteExplorerFrame(visible: false, explorerClass: true);
        var browser = AddPanel(fixture, panel, AllItems, out var info);
        fixture.SetExplorerWindows(panel.Handle);
        fixture.SetShownExplorerWindows(panel.Handle);
        Check.That(!await (Task<bool>)fixture.Invoke("TryRestoreNewExplorerWindowAsync", browser, info, false)!,
            "Control Panel must be rejected before the automatic restore owner adopts the window.");
        Check.Equal(0, FieldCount(fixture, "_sessionRestoreExclusions"),
            "An unrelated Control Panel launch must not consume an Explorer restore opportunity.");
    });

    private static Task DoesNotCaptureControlPanel() => ExplorerSessionNativeTests.WithCapture((fixture, _) =>
    {
        using var panel = new RemoteExplorerFrame(visible: false, explorerClass: true);
        AddPanel(fixture, panel, AllItems);
        fixture.SetExplorerWindows(panel.Handle);
        fixture.SetShownExplorerWindows(panel.Handle);
        fixture.Invoke("CaptureExplorerSessions");
        Check.Equal(0, FieldCount(fixture, "_sessionWindows"),
            "Closing Control Panel must never replace the saved folder group.");
        return Task.CompletedTask;
    });

    private static Task RecognizesNamespaceVariants()
    {
        foreach (var location in ControlPanelLocations)
        {
            foreach (var variant in new[] { location, location + "\\", " \"" + location.ToLowerInvariant() + "\" ",
                         location.Replace('\\', '/') })
                Check.That(ExplorerLocationPolicy.IsControlPanel(variant), "Control Panel variant was not excluded: " + variant);
            if (location.StartsWith("shell:::"))
            {
                Check.That(ExplorerLocationPolicy.IsControlPanel(location[6..]), "Raw shell parsing names must be excluded.");
                Check.That(ExplorerLocationPolicy.IsControlPanel(location[8..]), "Bare CLSID parsing names must be excluded.");
            }
        }
        return Task.CompletedTask;
    }

    private static Task PreservesOrdinaryLocations()
    {
        foreach (var location in new string?[] { null, "", " ", ThisPc,
                     "shell:::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}",
                     "shell:::{645FF040-5081-101B-9F08-00AA002F954E}",
                     @"C:\Control Panel", @"C:\shell:ControlPanelFolder", @"\\server\ControlPanelFolder",
                     "file:///C:/ControlPanelFolder", @"C:\::{21EC2020-3AEA-1069-A2DD-08002B30309D}",
                     Category + "Backup", "shell:ControlPanelFolderBackup" })
            Check.That(!ExplorerLocationPolicy.IsControlPanel(location), "An ordinary location was excluded: " + location);
        return Task.CompletedTask;
    }

    private static Task SourcesBypassAutomaticRestore() => ExplorerSessionNativeTests.WithCapture(async (fixture, _) =>
    {
        using var panel = new RemoteExplorerFrame(visible: false, explorerClass: true);
        var browser = AddPanel(fixture, panel, AllItems, out var info);
        fixture.SetExplorerWindows(panel.Handle);
        fixture.SetShownExplorerWindows(panel.Handle);
        fixture.EnableMerging();
        await (Task)fixture.Invoke("ProcessRegisteredShellWindowAsync", browser, info)!;
        Check.Equal(0, FieldCount(fixture, "_sessionRestoreExclusions"),
            "Registration must release Control Panel before the restore entry point consumes its launch.");
        Check.Equal(0, fixture.MergeSourceCount, "A Control Panel source must not be concealed for merging.");
        Check.That(info.EventsHooked, "Read-only navigation tracking must remain attached to the independent window.");
    });

    private static Task RejectsTabCreation() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var browser = fixture.AddBrowser(out var info, fixture.Window.FirstTab, locationUrl: () => AllItems);
        info.Location = AllItems;
        foreach (var classicOnly in new[] { false, true })
        {
            fixture.DirectTabUnsupported = classicOnly;
            await ExpectCancellation((Task)fixture.Invoke("CreateTabAtLocationAsync", fixture.Window.Handle,
                info.Identity, new[] { fixture.Window.FirstTab }, Fixture.Location)!,
                "Direct creation must cancel, not authorize a classic fallback into Control Panel.");
        }
        await ExpectCancellation((Task)fixture.Invoke("RequestToOpenNewTab", fixture.Window.Handle, false, true)!,
            "The classic new-tab command must also reject Control Panel.");
    }, tabCount: 1);

    private static Task DetectsLateNavigation() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        fixture.AddBrowser(out var info, fixture.Window.FirstTab, locationUrl: () => AllItems);
        Check.That(!ExplorerLocationPolicy.IsControlPanel(info.Location), "The fixture must begin with stale folder metadata.");
        await ExpectCancellation((Task)fixture.Invoke("CreateTabAtLocationAsync", fixture.Window.Handle,
            info.Identity, new[] { fixture.Window.FirstTab }, Fixture.Location)!,
            "A queued navigation event must not allow BrowseObject to alter the live Control Panel view.");
        Check.That(ExplorerLocationPolicy.IsControlPanel(info.Location), "The live exclusion must refresh the tab metadata.");
        Check.That(!fixture.DirectTabUnsupported, "An excluded window does not mean direct tabs are unsupported.");
    }, tabCount: 1);

    private static Task RejectsOpenRequest() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        // A cancelled lifetime guarantees that a regression stops at the lock, never at a desktop call.
        fixture.HookLifetime.Cancel();
        Check.That(!await (Task<bool>)fixture.Invoke("OpenTabNavigateWithSelection", new WindowRecord(AllItems),
            fixture.Window.Handle)!, "An excluded request must be rejected before any merge work starts.");
    }, tabCount: 1);

    private static Task TracksNavigationChanges() => ExplorerSessionNativeTests.WithCapture((fixture, _) =>
    {
        using var frame = new RemoteExplorerFrame(visible: false, explorerClass: true);
        var browser = fixture.AddBrowser(out var info, frame.Tab, handle: frame.Handle);
        fixture.Invoke("HookWindowEvents", browser, info);
        fixture.SetExplorerWindows(frame.Handle);
        fixture.SetShownExplorerWindows(frame.Handle);
        fixture.Invoke("CaptureExplorerSessions");
        Check.Equal(1, FieldCount(fixture, "_sessionWindows"), "The normal folder must be captured before navigation.");
        var navigation = (Delegate)typeof(WindowInfo).GetProperty("OnNavigateHandler")!.GetValue(info)!;
        navigation.DynamicInvoke(new object?[] { browser, AllItems });
        Check.That((bool)fixture.Invoke("IsControlPanelWindow", frame.Handle)!, "Navigation must immediately change window classification.");
        Check.Equal(0, FieldCount(fixture, "_sessionWindows"), "Control Panel must invalidate the previous folder snapshot before close.");
        Check.That(!(bool)fixture.Invoke("HasKnownExplorerLocation", frame.Handle)!, "The navigated frame must no longer be a tab target.");
        navigation.DynamicInvoke(new object?[] { browser, "file:///C:/WinTab-lifetime" });
        Check.That(!(bool)fixture.Invoke("IsControlPanelWindow", frame.Handle)!, "A later user navigation must not retain a stale exclusion.");
        Check.That((bool)fixture.Invoke("HasKnownExplorerLocation", frame.Handle)!, "The user's normal folder becomes eligible again.");
        return Task.CompletedTask;
    });

    private static Task DoesNotBlockFolderRestore() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var panel = new RemoteExplorerFrame(visible: false, explorerClass: true);
        using var folder = new RemoteExplorerFrame(visible: false, explorerClass: true);
        AddPanel(fixture, panel, Category);
        fixture.AddBrowser(out _, folder.Tab, handle: folder.Handle);
        fixture.SetExplorerWindows(folder.Handle, panel.Handle);
        fixture.SetShownExplorerWindows(folder.Handle, panel.Handle);
        Check.That(!(bool)fixture.Invoke("HasOtherShownSessionWindow", folder.Handle)!,
            "An open Control Panel must not count as a second folder window and suppress normal restoration.");
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task ChecksDoNotWaitForCatalog() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        fixture.AddBrowser(out var info, fixture.Window.FirstTab);
        info.Location = AllItems;
        var gate = typeof(ExplorerWatcher).GetField("_windowEntryDictLock", BindingFlags.Instance | BindingFlags.NonPublic)!
            .GetValue(fixture.Watcher)!;
        using var release = new ManualResetEventSlim();
        var held = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var owner = Task.Run(() =>
        {
            lock (gate)
            {
                held.SetResult();
                release.Wait();
            }
        });
        await held.Task;
        try
        {
            var check = Task.Run(() => (bool)fixture.Invoke("IsControlPanelWindow", fixture.Window.Handle)!);
            Check.That(await check.WaitAsync(TimeSpan.FromSeconds(1)),
                "WinEvent exclusion must remain readable while the STA catalog owner waits for Explorer COM.");
        }
        finally
        {
            release.Set();
            await owner;
        }
    }, tabCount: 1);

    private static async Task ExpectCancellation(Task task, string message)
    {
        try { await task; }
        catch (OperationCanceledException) { return; }
        throw new InvalidOperationException(message);
    }

    private static object AddPanel(Fixture fixture, RemoteExplorerFrame panel, string location) =>
        AddPanel(fixture, panel, location, out _);

    private static object AddPanel(Fixture fixture, RemoteExplorerFrame panel, string location, out WindowInfo info)
    {
        var browser = fixture.AddBrowser(out info, panel.Tab, handle: panel.Handle, locationUrl: () => location);
        info.Location = location;
        fixture.Invoke("HookWindowEvents", browser, info);
        return browser;
    }

    private static int FieldCount(Fixture fixture, string name) =>
        ((ICollection)typeof(ExplorerWatcher).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!
            .GetValue(fixture.Watcher)!).Count;
}
