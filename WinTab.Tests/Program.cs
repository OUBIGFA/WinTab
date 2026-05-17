using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.Interop;
using WinTab.WinAPI;
using ShellServiceProvider = WinTab.Interop.IServiceProvider;

if (args.Length > 0 && StringComparer.OrdinalIgnoreCase.Equals(args[0], "--stress"))
    return await ExplorerStressTest.RunAsync(args);
if (args.Length > 0 && StringComparer.OrdinalIgnoreCase.Equals(args[0], "--activation-stress"))
    return await ExplorerStressTest.RunActivationAsync(args);
if (args.Length > 0 && StringComparer.OrdinalIgnoreCase.Equals(args[0], "--recovery-stress"))
    return await ExplorerStressTest.RunRecoveryAsync(args);
if (args.Length > 0 && StringComparer.OrdinalIgnoreCase.Equals(args[0], "--reuse-stress"))
    return await ExplorerStressTest.RunReuseAsync(args);
if (args.Length > 0 && StringComparer.OrdinalIgnoreCase.Equals(args[0], "--default-location-stress"))
    return await ExplorerStressTest.RunDefaultLocationAsync(args);
if (args.Length > 0 && StringComparer.OrdinalIgnoreCase.Equals(args[0], "--user-default-stress"))
    return await ExplorerStressTest.RunUserDefaultAsync(args);
if (args.Length > 0 && StringComparer.OrdinalIgnoreCase.Equals(args[0], "--mixed-default-folder-stress"))
    return await ExplorerStressTest.RunMixedDefaultFolderAsync(args);

return await ExplorerLaunchLocationResolverTests.RunAll();

internal static class ExplorerLaunchLocationResolverTests
{
    public static async Task<int> RunAll()
    {
        var tests = new (string Name, Func<Task> Body)[]
        {
            ("waits for the real folder when Explorer first reports This PC", WaitsForRealFolderAfterTransientDefault),
            ("waits longer for delayed external Shell folder launches", WaitsForDelayedExternalShellFolderAfterDefault),
            ("returns the default folder only after the startup location stays default", ReturnsDefaultAfterTimeout),
            ("busy startup locations wait for the real folder", BusyStartupLocationWaitsForRealFolder),
            ("waits for a non-default location to stabilize", WaitsForStableNonDefaultLocation),
            ("normalizes file URLs to local filesystem paths", NormalizesFileUrlsToLocalPaths),
            ("keeps web URLs usable", KeepsWebUrlsUsable),
            ("SelectTabByHandle uses fast selection before ExplorerTabUtility fallback", ExplorerTabSelectionTests.SelectTabByHandleUsesFastSelectionBeforeExplorerTabUtilityFallback),
            ("tab merge uses the ExplorerTabUtility fast path", ExplorerTabSelectionTests.TabMergeUsesExplorerTabUtilityFastPath),
            ("window registration batches do not serialize slow location resolution", ExplorerTabSelectionTests.WindowRegistrationBatchesDoNotSerializeSlowLocationResolution),
            ("tab navigation failure is closed on a short bounded path", ExplorerTabSelectionTests.TabNavigationFailureIsShortBoundedPath),
            ("tab reuse avoids fixed waits and reuses automation root", ExplorerTabSelectionTests.TabReuseAvoidsFixedWaitsAndReusesAutomationRoot),
            ("failed navigation closes the transient This PC tab", ExplorerTabSelectionTests.FailedNavigationClosesTransientThisPcTab),
            ("WinEvent hook hides merge source windows without residuals", ExplorerTabSelectionTests.WinEventHookHidesMergeSourceWindowsWithoutResiduals),
            ("merge source conceal pulse is bounded and event-triggered", ExplorerTabSelectionTests.MergeSourceConcealPulseIsBoundedAndEventTriggered),
            ("ShellWindows adoption hides merge sources before tracking them", ExplorerTabSelectionTests.ShellWindowsAdoptionHidesMergeSourcesBeforeTracking),
            ("ShellWindows registration callback stays non-blocking", ExplorerTabSelectionTests.ShellWindowRegistrationCallbackStaysNonBlocking),
            ("stable default-location Explorer windows remain mergeable without residuals", ExplorerTabSelectionTests.StableDefaultLocationWindowsRemainMergeableWithoutResiduals),
            ("folder merges avoid default-only Explorer targets when a real folder target exists", ExplorerTabSelectionTests.FolderMergesAvoidDefaultOnlyTargets),
            ("mergeable default-location intermediate closes only after target tab succeeds", ExplorerTabSelectionTests.MergeableDefaultLocationIntermediateClosesAfterTargetTabSucceeds),
            ("released merge sources are protected from late re-hide events", ExplorerTabSelectionTests.ReleasedMergeSourcesAreProtectedFromLateRehideEvents),
            ("HideWindow reapplies transparency after Explorer resets styles", ExplorerTabSelectionTests.HideWindowReappliesTransparency),
            ("hidden Explorer windows are restored on lifecycle boundaries", ExplorerTabSelectionTests.HiddenExplorerWindowsAreRestoredOnLifecycleBoundaries),
            ("orphaned transparent Explorer windows are recovered without cache", ExplorerTabSelectionTests.OrphanedTransparentExplorerWindowsAreRecoveredWithoutCache),
            ("Shell initialization skips duplicate ShellWindows entries", ExplorerTabSelectionTests.ShellInitializationSkipsDuplicateShellWindowsEntries),
            ("tab reuse excludes only the merge source", ExplorerTabSelectionTests.TabReuseExcludesMergeSource),
            ("tab reuse foregrounds the target window before selecting", ExplorerTabSelectionTests.TabReuseForegroundsTargetBeforeSelection),
            ("tab reuse never creates a duplicate after finding an existing path", ExplorerTabSelectionTests.TabReuseDoesNotDuplicateAfterSelectionFailure),
            ("WindowRegistered restores pre-hidden windows that are not merged", ExplorerTabSelectionTests.WindowRegisteredRestoresPreHiddenUnmergedWindows),
            ("merge source close is verified before hidden tracking is cleared", ExplorerTabSelectionTests.MergeSourceCloseIsVerifiedBeforeHiddenTrackingIsCleared),
            ("Early WinEvent hide does not wait on merge target discovery", ExplorerTabSelectionTests.EarlyHideDoesNotWaitOnMergeTargetDiscovery),
            ("first run settings window fits without scrolling", ExplorerTabSelectionTests.FirstRunSettingsWindowFitsWithoutScrolling),
            ("settings resize writes are debounced and flushed", ExplorerTabSelectionTests.SettingsResizeWritesAreDebouncedAndFlushed),
            ("Explorer hot paths avoid repeated polling work", ExplorerTabSelectionTests.ExplorerHotPathsAvoidRepeatedPollingWork),
            ("window merge uses a single registration lifecycle", ExplorerTabSelectionTests.WindowMergeUsesSingleRegistrationLifecycle),
            ("update checks prefer the current architecture installer", ExplorerTabSelectionTests.UpdateCheckPrefersCurrentArchitectureInstaller)
        };

        var failed = 0;
        foreach (var (name, body) in tests)
        {
            try
            {
                await body();
                Console.WriteLine($"PASS {name}");
            }
            catch (Exception ex)
            {
                failed++;
                Console.Error.WriteLine($"FAIL {name}: {ex.Message}");
            }
        }

        return failed == 0 ? 0 : 1;
    }

    private static async Task WaitsForRealFolderAfterTransientDefault()
    {
        var samples = new Queue<string>([
            KnownLocations.ThisPc,
            KnownLocations.ThisPc,
            KnownLocations.TargetFolder,
            KnownLocations.TargetFolder
        ]);

        var resolver = CreateFastResolver();
        var resolved = await resolver.ResolveAsync(
            () => samples.Count > 0 ? samples.Dequeue() : KnownLocations.TargetFolder,
            IsDefaultLocation);

        AssertEqual(KnownLocations.TargetFolder, resolved);
    }

    private static async Task ReturnsDefaultAfterTimeout()
    {
        var resolver = CreateFastResolver();
        var resolved = await resolver.ResolveAsync(() => KnownLocations.ThisPc, IsDefaultLocation);

        AssertEqual(KnownLocations.ThisPc, resolved);
    }

    private static async Task WaitsForDelayedExternalShellFolderAfterDefault()
    {
        var start = Environment.TickCount64;
        var resolver = CreateFastResolver();
        var releasedStartupLocation = false;

        var resolved = await resolver.ResolveAsync(
            () => Environment.TickCount64 - start < 95 ? KnownLocations.ThisPc : KnownLocations.TargetFolder,
            IsDefaultLocation,
            onStartupLocationRetained: _ =>
            {
                releasedStartupLocation = true;
                return Task.CompletedTask;
            });

        Assert(releasedStartupLocation,
            "A merge source that still reports This PC after the short wait must be released instead of kept hidden.");
        AssertEqual(KnownLocations.TargetFolder, resolved);
    }

    private static async Task WaitsForStableNonDefaultLocation()
    {
        var samples = new Queue<string>([
            KnownLocations.ThisPc,
            KnownLocations.Downloads,
            KnownLocations.TargetFolder,
            KnownLocations.TargetFolder
        ]);

        var resolver = CreateFastResolver();
        var resolved = await resolver.ResolveAsync(
            () => samples.Count > 0 ? samples.Dequeue() : KnownLocations.TargetFolder,
            IsDefaultLocation);

        AssertEqual(KnownLocations.TargetFolder, resolved);
    }

    private static async Task BusyStartupLocationWaitsForRealFolder()
    {
        var start = Environment.TickCount64;
        var resolver = CreateFastResolver();

        var resolved = await resolver.ResolveAsync(
            () => Environment.TickCount64 - start < 95 ? KnownLocations.ThisPc : KnownLocations.TargetFolder,
            IsDefaultLocation,
            isBusy: () => Environment.TickCount64 - start < 95);

        AssertEqual(KnownLocations.TargetFolder, resolved);
    }

    private static Task NormalizesFileUrlsToLocalPaths()
    {
        var normalized = Helper.NormalizeLocation("file:///C:/Users/BIGFA/Downloads");

        AssertEqual(@"C:\Users\BIGFA\Downloads", normalized);
        return Task.CompletedTask;
    }

    private static Task KeepsWebUrlsUsable()
    {
        var normalized = Helper.NormalizeLocation("https://example.com/path/to/file");

        AssertEqual("https://example.com/path/to/file", normalized);
        return Task.CompletedTask;
    }

    private static ExplorerLaunchLocationResolver CreateFastResolver()
    {
        return new ExplorerLaunchLocationResolver(new ExplorerLaunchLocationResolver.Options(
            DefaultLocationWaitMs: 80,
            StableLocationWaitMs: 15,
            PollIntervalMs: 1,
            MaximumStartupLocationWaitMs: 160));
    }

    private static bool IsDefaultLocation(string location)
    {
        return StringComparer.OrdinalIgnoreCase.Equals(location, KnownLocations.ThisPc);
    }

    private static void AssertEqual(string expected, string actual)
    {
        if (!StringComparer.OrdinalIgnoreCase.Equals(expected, actual))
            throw new InvalidOperationException($"expected '{expected}', got '{actual}'");
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }

    private static class KnownLocations
    {
        public const string ThisPc = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
        public const string Downloads = "file:///C:/Users/BIGFA/Downloads";
        public const string TargetFolder = "file:///E:/WinTabStress/Target";
    }
}

internal static class ExplorerTabSelectionTests
{
    public static Task SelectTabByHandleUsesFastSelectionBeforeExplorerTabUtilityFallback()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "public async Task SelectTabByHandle") +
                         ExtractMethodBody(source, "private async Task<bool> TrySelectTabByHandleDirectAsync") +
                         ExtractMethodBody(source, "private async Task<bool> SelectTabByUniqueNameVerified") +
                         ExtractMethodBody(source, "private async Task<bool> TrySelectTabByCyclingAsync");

        var nameSelectIndex = methodBody.IndexOf("TrySelectSingleTabByAutomationName(windowHandle, tabName)", StringComparison.Ordinal);
        var fallbackIndex = methodBody.IndexOf("TrySelectTabByCyclingAsync(windowHandle, tabHandle, timeoutMs)", StringComparison.Ordinal);
        Assert(nameSelectIndex >= 0 && fallbackIndex > nameSelectIndex,
            "Reuse should try WinTab's direct UI Automation selection before falling back to ExplorerTabUtility-style cycling.");
        Assert(methodBody.Contains("InvalidateAutomationRoot(windowHandle)", StringComparison.Ordinal),
            "Selection should refresh stale UI Automation state before using the slower cycling fallback.");
        Assert(methodBody.Contains("SelectTabByIndex(windowHandle, i)", StringComparison.Ordinal) &&
               methodBody.Contains("i < tabs.Length", StringComparison.Ordinal),
            "ExplorerTabUtility's cycling fallback must remain available so reuse still works when direct selection fails.");

        return Task.CompletedTask;
    }

    public static Task TabMergeUsesExplorerTabUtilityFastPath()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection") +
                         ExtractMethodBody(source, "private async Task<bool> NavigateNewTabToTargetAsync");

        Assert(methodBody.Contains("SelectTabByUniqueNameVerified(windowHandle, existingTab, 700, existingWindow)", StringComparison.Ordinal),
            "Tab reuse should follow ExplorerTabUtility's direct handle activation path.");
        Assert(methodBody.Contains("ListenForNewExplorerTabAsync(mainWindowHWnd, currentTabs, 2_000)", StringComparison.Ordinal),
            "New tab creation should use the short ExplorerTabUtility wait window.");
        Assert(methodBody.Contains("FindShellWindowByTabHandle(newTabHandle, mainWindowHWnd)", StringComparison.Ordinal) &&
               methodBody.Contains("2_000", StringComparison.Ordinal),
            "New tab navigation should bind the ShellWindows object by tab handle without adding slow address-bar fallbacks.");
        Assert(!source.Contains("NavigateActiveTabByAddressBar", StringComparison.Ordinal) &&
               !source.Contains("TryPasteAddressText", StringComparison.Ordinal) &&
               !source.Contains("KeyboardSimulator.SendText", StringComparison.Ordinal),
            "Opening a tab must not drive Explorer through the address bar or clipboard.");
        var registrationBody = ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");
        var handleIndex = registrationBody.IndexOf("GetTabHandle(window)", StringComparison.Ordinal);
        var hideIndex = registrationBody.IndexOf("HideMergeSourceWindow(hWnd)", StringComparison.Ordinal);
        var tabCountIndex = registrationBody.IndexOf("WaitForExplorerTabCount(hWnd)", StringComparison.Ordinal);
        Assert(hideIndex >= 0 && handleIndex > hideIndex && tabCountIndex > handleIndex,
            "WindowRegistered must hide the source first, then bind the ShellWindow to its tab handle before slower checks.");
        Assert(!methodBody.Contains("3_500", StringComparison.Ordinal),
            "Opening a new target tab must not wait several seconds before sending navigation.");
        Assert(!methodBody.Contains("restore-previous", StringComparison.OrdinalIgnoreCase),
            "The merge path must not restore previous tabs while opening a target tab.");
        Assert(!methodBody.Contains("CloseUnexpectedDefaultTabs", StringComparison.Ordinal),
            "The merge path must not run a broad default-tab cleanup pass.");
        Assert(!methodBody.Contains("newTabTimeoutMs", StringComparison.Ordinal),
            "The merge path must not carry the old long timeout path.");
        Assert(!source.Contains("SelectTabByKnownIndexVerified", StringComparison.Ordinal),
            "The known-index merge path should be removed, not bypassed.");
        var lockReleaseIndex = methodBody.IndexOf("_toOpenWindowsLock.Release()", StringComparison.Ordinal);
        var navigationIndex = methodBody.IndexOf("NavigateNewTabToTargetAsync(window, windowToOpen.Location)", StringComparison.Ordinal);
        Assert(lockReleaseIndex >= 0 && navigationIndex > lockReleaseIndex,
            "Navigation verification should run after releasing the new-tab creation lock so rapid folder merges can overlap navigation checks.");

        return Task.CompletedTask;
    }

    public static Task WindowRegistrationBatchesDoNotSerializeSlowLocationResolution()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowsAsync");

        Assert(methodBody.Contains("Task.WhenAll", StringComparison.Ordinal) &&
               methodBody.Contains("ProcessRegisteredShellWindowAsync", StringComparison.Ordinal),
            "Rapid Explorer opens must resolve source locations concurrently; one slow transient This PC window must not block the whole registration batch.");
        Assert(!methodBody.Contains("foreach (var (window, windowInfo) in windows)\r\n                await ProcessRegisteredShellWindowAsync(window, windowInfo);", StringComparison.Ordinal),
            "Registration processing must not await each hidden source window serially.");

        return Task.CompletedTask;
    }

    public static Task TabNavigationFailureIsShortBoundedPath()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var openBody = ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection");
        var navigateBody = ExtractMethodBody(source, "private async Task<bool> NavigateNewTabToTargetAsync");

        Assert(!openBody.Contains("for (var attempt = 1; attempt <= 3; attempt++)", StringComparison.Ordinal),
            "Opening a target tab must not retry the whole tab creation flow while the source window is hidden.");
        Assert(!navigateBody.Contains("for (var navigateAttempt = 1; navigateAttempt <= 3; navigateAttempt++)", StringComparison.Ordinal),
            "Navigation verification must not retry Navigate2 in a nested slow loop.");
        Assert(!navigateBody.Contains("WaitForNavigation(window, targetLocation, 5_000)", StringComparison.Ordinal),
            "A wrong This PC tab must be detected and closed on a short bounded verification path, not after a five-second wait.");
        Assert(navigateBody.Contains("NavigateComplete2", StringComparison.Ordinal) &&
               source.Contains("NavigationVerificationWaitMs = 1_200", StringComparison.Ordinal) &&
               navigateBody.Contains("WaitForNavigation(window, targetLocation, NavigationVerificationWaitMs)", StringComparison.Ordinal),
            "Navigation should follow ExplorerTabUtility's NavigateComplete signal, then keep WinTab's custom short target verification.");

        return Task.CompletedTask;
    }

    public static Task TabReuseAvoidsFixedWaitsAndReusesAutomationRoot()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var openBody = ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection");
        var selectBody = ExtractMethodBody(source, "private bool TrySelectSingleTabByAutomationName");

        Assert(!openBody.Contains("Task.Delay(60)", StringComparison.Ordinal),
            "Reuse should not pay a fixed sleep after foregrounding; direct selection is already verified by the active tab handle.");
        Assert(selectBody.Contains("GetAutomationRoot(windowHandle)", StringComparison.Ordinal),
            "Reuse selection should use the cached automation root instead of rebuilding UI Automation state for every tab activation.");
        Assert(!selectBody.Contains("AutomationElement.FromHandle(windowHandle)", StringComparison.Ordinal),
            "AutomationElement.FromHandle is cached elsewhere and should not run on every reuse selection.");

        return Task.CompletedTask;
    }

    public static Task FailedNavigationClosesTransientThisPcTab()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection") +
                         ExtractMethodBody(source, "private async Task<bool> NavigateNewTabToTargetAsync") +
                         ExtractMethodBody(source, "private bool AreLocationsEquivalent") +
                         ExtractMethodBody(source, "private nint GetMainWindowHWnd") +
                         ExtractMethodBody(source, "private bool IsStableMergeTargetWindow") +
                         ExtractMethodBody(source, "private static bool IsFallbackMergeTargetWindow") +
                         ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");

        Assert(methodBody.Contains("WaitForNavigation(window, targetLocation, NavigationVerificationWaitMs)", StringComparison.Ordinal) &&
               source.Contains("NavigationVerificationWaitMs = 1_200", StringComparison.Ordinal) &&
               methodBody.Contains("AreLocationsEquivalent(TryGetLocation(window), targetLocation)", StringComparison.Ordinal),
            "The newly opened tab must be checked against the requested target folder.");
        Assert(methodBody.Contains("CloseFailedNewTabAsync(mainWindowHWnd, newTabHandle)", StringComparison.Ordinal),
            "A tab that remains at This PC or another wrong location must be closed instead of left behind.");
        Assert(methodBody.Contains("Helper.IsFileExplorerForeground(out var foregroundWindow)", StringComparison.Ordinal),
            "The merge target must prefer the current foreground Explorer instead of a stale background This PC window.");
        Assert(methodBody.Contains("IsPreferredMergeTargetWindow(foregroundWindow, otherThan, targetLocation)", StringComparison.Ordinal),
            "The foreground window must be stable and preferred for the resolved target, not another transient or default-only source window.");
        Assert(methodBody.Contains("GetMainWindowHWnd(hWnd)", StringComparison.Ordinal) ||
               methodBody.Contains("GetMainWindowHWnd(windowToOpen.Handle)", StringComparison.Ordinal),
            "Auto-merge must resolve the target Explorer while excluding the transient source window.");
        Assert(methodBody.Contains("IsFallbackMergeTargetWindow(h, otherThan)", StringComparison.Ordinal),
            "If the strict cached target is unavailable, merge should fall back to ExplorerTabUtility's tagged-window selection instead of opening a new window.");
        Assert(methodBody.Contains("Helper.NormalizeLocation(location)", StringComparison.Ordinal) &&
               methodBody.Contains("Helper.NormalizeLocation(targetLocation)", StringComparison.Ordinal),
            "Navigation completion checks must normalize local paths and file URLs before comparing them.");

        return Task.CompletedTask;
    }

    public static Task WinEventHookHidesMergeSourceWindowsWithoutResiduals()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "private bool TryHideIncomingExplorerWindow") +
                         ExtractMethodBody(source, "private void TryHideRegisteredMergeSourceWindow") +
                         ExtractMethodBody(source, "private void OnWindowShown") +
                         ExtractMethodBody(source, "private void OnShellWindowRegistered") +
                         ExtractMethodBody(source, "private List<(InternetExplorer Window, WindowInfo WindowInfo)> AdoptNewShellWindows") +
                         ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync") +
                         ExtractMethodBody(source, "private void HideMergeSourceWindow") +
                         ExtractMethodBody(source, "private async Task RestoreMergeSourceWindowAsync") +
                         ExtractMethodBody(source, "private void RemoveMergeSourceTracking") +
                         ExtractMethodBody(source, "private void StartMergeSourceConcealPulse") +
                         ExtractMethodBody(source, "private void ConcealMergeSourceWindowsOnce") +
                         ExtractMethodBody(source, "private void InitializeShellObjects") +
                         ExtractMethodBody(source, "private sealed class WinEventHookThread");

        Assert(methodBody.Contains("SetWinEventHook(WinApi.EVENT_SYSTEM_FOREGROUND, WinApi.EVENT_SYSTEM_FOREGROUND", StringComparison.Ordinal) &&
               methodBody.Contains("SetWinEventHook(WinApi.EVENT_OBJECT_CREATE, WinApi.EVENT_OBJECT_SHOW", StringComparison.Ordinal),
            "WinEvents should wake ShellWindows registration and hide mergeable source windows early.");
        Assert(methodBody.Contains("ScheduleShellWindowRegistration(1)", StringComparison.Ordinal),
            "Foreground/create/show events should trigger quick registration processing.");
        Assert(methodBody.Contains("TryHideIncomingExplorerWindow(hWnd)", StringComparison.Ordinal),
            "The WinEvent path must conceal mergeable source windows before Explorer paints a visible flash.");
        Assert(source.Contains("GetExplorerTopLevelWindow(hWnd)", StringComparison.Ordinal) &&
               source.Contains("WinApi.GetAncestor(hWnd, WinApi.GA_ROOT)", StringComparison.Ordinal),
            "Early hide must normalize child WinEvent handles to the real Explorer top-level window.");
        Assert(methodBody.Contains("StartMergeSourceConcealPulse()", StringComparison.Ordinal) &&
               methodBody.Contains("ConcealMergeSourceWindowsOnce()", StringComparison.Ordinal),
            "Rapid Explorer bursts need a short event-triggered conceal pulse so late WinEvents do not produce visible flashes.");
        Assert(!methodBody.Contains("HasMergeTargetForEarlyConceal(hWnd)", StringComparison.Ordinal),
            "Early hide must not wait for merge-target discovery; registration later decides whether to merge or quickly release a user-opened Explorer window.");
        Assert(methodBody.Contains("_hookedTopLevelUseCounts.ContainsKey(hWnd)", StringComparison.Ordinal) &&
               !methodBody.Contains("HasHookedShellWindowForTopLevel(hWnd)", StringComparison.Ordinal),
            "Early hide must use the lightweight hooked-target cache instead of dictionary/COM target checks.");
        Assert(methodBody.Contains("GetAllExplorerTabs(hWnd).Take(2).Count() > 1", StringComparison.Ordinal),
            "Early hide must not conceal an existing multi-tab target Explorer window.");
        Assert(methodBody.Contains("TryHideRegisteredMergeSourceWindow(hWnd)", StringComparison.Ordinal),
            "A ShellWindows registration batch must hide every mergeable source before slow per-window processing.");
        Assert(methodBody.Contains("wasTrackedTopLevel") && methodBody.Contains("if (!wasTrackedTopLevel)"),
            "ShellWindows registrations for new tabs inside the target window must not hide the target top-level window.");
        Assert(methodBody.Contains("_processedHWnds.ContainsKey(hWnd)", StringComparison.Ordinal),
            "Registered hide must ignore handles that were already released after a successful merge.");
        Assert(methodBody.Contains("HasOtherTrackedShellWindowForTopLevel(window, hWnd)", StringComparison.Ordinal),
            "Processing must treat ShellWindows that share an adopted top-level window as tabs, not merge-source windows.");
        Assert(methodBody.Contains("PublishTabHandleForTrackedShellWindowAsync(window)", StringComparison.Ordinal),
            "Tab ShellWindows must publish their tab handle before the opener waits for the new tab binding.");
        Assert(methodBody.Contains("HideMergeSourceWindow(hWnd)", StringComparison.Ordinal) &&
               methodBody.Contains("_mergeSourceHWnds.TryAdd(hWnd, 0)", StringComparison.Ordinal),
            "Hidden windows must be explicitly owned as merge sources.");
        Assert(methodBody.Contains("CloseMergedSourceWindowAsync(window, hWnd)", StringComparison.Ordinal) &&
               source.Contains("WinApi.WM_CLOSE", StringComparison.Ordinal) &&
               source.Contains("window.Quit()", StringComparison.Ordinal),
            "The merge source Explorer window must be closed directly after a successful merge without blocking the merge queue.");
        Assert(methodBody.Contains("RemoveMergeSourceTracking(hWnd)", StringComparison.Ordinal) &&
               methodBody.Contains("RestoreMergeSourceWindowAsync(hWnd)", StringComparison.Ordinal),
            "Merge source ownership must be cleared on success and restored on failure.");
        Assert(methodBody.Contains("PreventWindowHiding(hWnd)", StringComparison.Ordinal),
            "A source that was closed or released must be protected from late show events re-hiding its handle.");
        Assert(!source.Contains("ScheduleStaleConcealedWindow", StringComparison.Ordinal) &&
               !source.Contains("CloseConcealedExplorerWindow", StringComparison.Ordinal),
            "Hidden source cleanup must be tied to the merge lifecycle, not stale fallback closers.");
        Assert(!source.Contains("RunConcealWatchdog", StringComparison.Ordinal) &&
               !source.Contains("RunConcealSweepAsync", StringComparison.Ordinal),
            "The merge path should not depend on custom high-frequency window sweeps.");

        return Task.CompletedTask;
    }

    public static Task MergeSourceConcealPulseIsBoundedAndEventTriggered()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var startBody = ExtractMethodBody(source, "public void StartHook");
        var pulseBody = ExtractMethodBody(source, "private void StartMergeSourceConcealPulse");
        var windowShownBody = ExtractMethodBody(source, "private void OnWindowShown");
        var registeredBody = ExtractMethodBody(source, "private void OnShellWindowRegistered");

        Assert(!source.Contains("StartMergeSourceConcealPulse(0)", StringComparison.Ordinal),
            "The conceal pulse must not run forever after startup; that keeps scanning Explorer windows and can re-hide later user-opened This PC windows.");
        Assert(startBody.Contains("StartMergeSourceConcealPulse(500)", StringComparison.Ordinal),
            "Startup should only run a short warm-up conceal pulse.");
        Assert(source.Contains("private void StartMergeSourceConcealPulse(int durationMs = 1_200)", StringComparison.Ordinal),
            "Event-triggered conceal should stay bounded so hidden merge-source protection cannot linger for seconds after activity stops.");
        Assert(windowShownBody.Contains("StartMergeSourceConcealPulse()", StringComparison.Ordinal) &&
               registeredBody.Contains("StartMergeSourceConcealPulse()", StringComparison.Ordinal),
            "WinEvent and ShellWindows registration events should still trigger the short conceal pulse for rapid Explorer opens.");
        Assert(pulseBody.Contains("DateTime.UtcNow.AddMilliseconds(Math.Max(1, durationMs))", StringComparison.Ordinal) &&
               !pulseBody.Contains("DateTime.MaxValue", StringComparison.Ordinal),
            "The pulse worker should be duration-bound instead of using an unbounded normal scanning window.");

        return Task.CompletedTask;
    }

    public static Task ShellWindowRegistrationCallbackStaysNonBlocking()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "private void OnShellWindowRegistered");
        var scheduleBody = ExtractMethodBody(source, "private void ScheduleShellWindowRegistration");

        Assert(methodBody.Contains("ScheduleShellWindowRegistration(1)", StringComparison.Ordinal),
            "ShellWindows registration should schedule merge work after returning to Explorer.");
        Assert(!methodBody.Contains("AdoptNewShellWindows", StringComparison.Ordinal),
            "ShellWindows registration must not synchronously enumerate ShellWindows from Explorer's COM event callback.");
        Assert(scheduleBody.Contains("HasUntrackedShellWindows()", StringComparison.Ordinal),
            "The registration scheduler must rescan after clearing its in-flight flag so rapid windows cannot be dropped.");

        return Task.CompletedTask;
    }

    public static Task ShellWindowsAdoptionHidesMergeSourcesBeforeTracking()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "private List<(InternetExplorer Window, WindowInfo WindowInfo)> AdoptNewShellWindows");

        var hideIndex = methodBody.IndexOf("HideMergeSourceWindow(hWnd)", StringComparison.Ordinal);
        var addIndex = methodBody.IndexOf("_windowEntryDict.Add(window, windowInfo)", StringComparison.Ordinal);
        Assert(hideIndex >= 0 && addIndex > hideIndex,
            "A new top-level Explorer merge source should be hidden before it is tracked, otherwise the conceal pulse treats it as an adopted target and lets it flash.");
        Assert(methodBody.Contains("singleTabTopLevelsInBatch", StringComparison.Ordinal) &&
               methodBody.Contains("tabCount <= 1", StringComparison.Ordinal) &&
               methodBody.Contains("wasTrackedTopLevel || !singleTabTopLevelsInBatch.Add(hWnd)", StringComparison.Ordinal),
            "Duplicate ShellWindows wrappers for the same single-tab top-level must be skipped so they do not block This PC merging as fake sibling tabs.");
        Assert(!methodBody.Contains("CloseRetainedStartupLocationIfMergeableAsync", StringComparison.Ordinal),
            "ShellWindows adoption must not close a retained This PC source before registration can decide whether it is a real user request or a busy folder launch.");

        return Task.CompletedTask;
    }

    public static Task StableDefaultLocationWindowsRemainMergeableWithoutResiduals()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");
        var normalizedBody = methodBody.Replace("\r\n", "\n", StringComparison.Ordinal);

        var resolveIndex = normalizedBody.IndexOf("ResolveInitialLocation(window, hWnd)", StringComparison.Ordinal);
        var tabHandleIndex = normalizedBody.IndexOf("GetTabHandle(window)", StringComparison.Ordinal);
        var releaseBranchStart = normalizedBody.IndexOf("if (string.IsNullOrWhiteSpace(location)", StringComparison.Ordinal);
        var sourceAliveIndex = normalizedBody.IndexOf("var sourceAlive = Helper.IsFileExplorerWindow(hWnd)", StringComparison.Ordinal);
        var targetSelectionIndex = normalizedBody.IndexOf("targetWindow = GetMainWindowHWnd(hWnd, location)", StringComparison.Ordinal);
        var openIndex = normalizedBody.IndexOf("OpenTabNavigateWithSelection(record, targetWindow)", StringComparison.Ordinal);
        Assert(resolveIndex >= 0 && releaseBranchStart > resolveIndex && targetSelectionIndex > releaseBranchStart && openIndex > targetSelectionIndex,
            "A stable This PC Explorer window must be allowed to reach target selection and the merge path after resolution.");
        Assert(tabHandleIndex > resolveIndex,
            "Startup location resolution must happen before tab-handle waits so This PC intermediates can be closed promptly.");
        Assert(normalizedBody.Contains("if (sourceAlive && !IsStartupExplorerLocation(location))\n            {", StringComparison.Ordinal),
            "Stable This PC intermediates must not wait on their soon-to-close source tab handle before the merge/reuse path can run.");
        var preResolveBody = normalizedBody[..resolveIndex];
        Assert(!preResolveBody.Contains("targetWindow != 0", StringComparison.Ordinal),
            "A stable This PC window must not be released before location resolution just because merge target discovery briefly lagged.");

        var releaseBranch = normalizedBody[releaseBranchStart..sourceAliveIndex];
        Assert(!releaseBranch.Contains("IsStartupExplorerLocation(location)", StringComparison.Ordinal),
            "A stable This PC location is a valid user-opened Explorer target and must not be released solely because it is This PC.");
        Assert(releaseBranch.Contains("string.IsNullOrWhiteSpace(location)", StringComparison.Ordinal) &&
               releaseBranch.Contains("shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}", StringComparison.Ordinal),
            "The early release branch should stay limited to unresolved locations and Control Panel, not stable This PC.");
        Assert(normalizedBody.Contains("if (sourceAlive && !IsStartupExplorerLocation(location))\n                HideMergeSourceWindow(hWnd);", StringComparison.Ordinal),
            "The merge path must avoid re-hiding a stable This PC source while still hiding real folder merge sources.");
        Assert(normalizedBody.Contains("CloseMergedSourceWindowAsync(window, hWnd)", StringComparison.Ordinal) &&
               normalizedBody.Contains("RemoveMergeSourceTracking(hWnd)", StringComparison.Ordinal),
            "A merged This PC source must be closed and removed from hidden tracking instead of remaining as a transparent residual window.");
        Assert(!normalizedBody.Contains("await OpenNewWindowWithSelection(record, lockToOpenWindows: false)", StringComparison.Ordinal),
            "A closed This PC intermediate must not be reopened as another top-level fallback that can become hidden in the background.");
        Assert(!methodBody.Contains("180_000", StringComparison.Ordinal),
            "Hidden merge candidates must not stay concealed for minutes.");

        return Task.CompletedTask;
    }

    public static Task FolderMergesAvoidDefaultOnlyTargets()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var registrationBody = ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");
        var targetBody = ExtractMethodBody(source, "private nint GetMainWindowHWnd") +
                         ExtractMethodBody(source, "private bool IsPreferredMergeTargetWindow") +
                         ExtractMethodBody(source, "private bool HasNonStartupShellWindowForTopLevel");

        Assert(registrationBody.Contains("GetMainWindowHWnd(hWnd, location)", StringComparison.Ordinal),
            "After resolving a real folder, merge target selection must use the resolved location to avoid a stale This PC target.");
        Assert(targetBody.Contains("ShouldPreferNonStartupTarget(targetLocation)", StringComparison.Ordinal) &&
               targetBody.Contains("HasNonStartupShellWindowForTopLevel", StringComparison.Ordinal),
            "Folder merges should prefer a target Explorer that already represents a real folder when one exists.");

        return Task.CompletedTask;
    }

    public static Task MergeableDefaultLocationIntermediateClosesAfterTargetTabSucceeds()
    {
        var resolverPath = FindRepoFile("WinTab", "Hooks", "ExplorerLaunchLocationResolver.cs");
        var watcherPath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var resolver = File.ReadAllText(resolverPath);
        var watcher = File.ReadAllText(watcherPath);
        var resolveBody = ExtractMethodBody(watcher, "private Task<string> ResolveInitialLocation");
        var restoreBody = ExtractMethodBody(watcher, "private static async Task RestoreHiddenExplorerWindowAsync");
        var registrationBody = ExtractMethodBody(watcher, "private async Task ProcessRegisteredShellWindowAsync");

        Assert(resolver.Contains("int DefaultLocationWaitMs = 60", StringComparison.Ordinal),
            "A non-mergeable stable This PC window should still be released quickly instead of staying transparent.");
        Assert(resolver.Contains("int MaximumStartupLocationWaitMs = 500", StringComparison.Ordinal),
            "Mergeable stable This PC windows should not wait the old multi-second startup window before merging.");
        Assert(resolver.Contains("int BusyStartupLocationWaitMs = 1_500", StringComparison.Ordinal) &&
               resolver.Contains("SafeIsBusy(isBusy)", StringComparison.Ordinal),
            "A startup-location source that is still busy should keep waiting for the real folder instead of being treated as This PC pollution.");
        Assert(resolveBody.Contains("isBusy: () => IsShellWindowBusy(window)", StringComparison.Ordinal),
            "Startup location resolution must use Explorer's Busy state before accepting This PC as the final target.");
        Assert(!watcher.Contains("CloseRetainedStartupLocationIfMergeableAsync", StringComparison.Ordinal),
            "The source This PC window must not be closed before the normal target-tab path succeeds; early close can swallow the user's Explorer open.");
        Assert(registrationBody.Contains("OpenTabNavigateWithSelection(record, targetWindow)", StringComparison.Ordinal) &&
               registrationBody.Contains("CloseMergedSourceWindowAsync(window, hWnd)", StringComparison.Ordinal),
            "Default-location sources should be closed by the same verified post-merge close path as folder sources.");
        Assert(restoreBody.Contains("50)", StringComparison.Ordinal),
            "Restoring a retained This PC window should poll quickly instead of waiting in coarse 200ms steps.");

        return Task.CompletedTask;
    }

    public static Task ReleasedMergeSourcesAreProtectedFromLateRehideEvents()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var restoreBody = ExtractMethodBody(source, "private async Task RestoreMergeSourceWindowAsync");

        Assert(restoreBody.Contains("PreventWindowHiding(hWnd)", StringComparison.Ordinal),
            "A released This PC/user Explorer window must be protected from late WinEvents that would hide it again.");

        return Task.CompletedTask;
    }

    public static Task HideWindowReappliesTransparency()
    {
        var sourcePath = FindRepoFile("WinTab", "Helpers", "Helper.cs");
        var source = File.ReadAllText(sourcePath);
        var hideBody = ExtractMethodBody(source, "public static void HideWindow");
        var showBody = ExtractMethodBody(source, "public static bool ShowWindow");

        var getOrAddIndex = hideBody.IndexOf("HiddenWindows.GetOrAdd", StringComparison.Ordinal);
        var alphaIndex = hideBody.LastIndexOf("SetLayeredWindowAttributes(hWnd, 0, 0, WinApi.LWA_ALPHA)", StringComparison.Ordinal);
        Assert(getOrAddIndex >= 0 && alphaIndex > getOrAddIndex,
            "HideWindow must reapply alpha=0 even when the window is already cached as hidden.");
        Assert(showBody.Contains("SetLayeredWindowAttributes(hWnd, 0, 255, WinApi.LWA_ALPHA)", StringComparison.Ordinal),
            "ShowWindow must restore alpha after a hidden window is released.");

        return Task.CompletedTask;
    }

    public static Task HiddenExplorerWindowsAreRestoredOnLifecycleBoundaries()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var initializeBody = ExtractMethodBody(source, "private void InitializeShellObjects");
        var stopBody = ExtractMethodBody(source, "public void StopHook");
        var disposeBody = ExtractMethodBody(source, "private void DisposeShellObjects");
        var registrationBody = ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");
        var removeBody = ExtractMethodBody(source, "private void RemoveWindowAndUnhookEvents");

        var startupRestoreIndex = initializeBody.IndexOf("RecoverHiddenExplorerWindows", StringComparison.Ordinal);
        var eventHookIndex = initializeBody.IndexOf("WindowRegistered +=", StringComparison.Ordinal);
        Assert(startupRestoreIndex >= 0 && eventHookIndex >= 0 && startupRestoreIndex < eventHookIndex,
            "Startup must recover orphaned hidden Explorer windows before registering new merge hooks.");
        Assert(stopBody.Contains("RecoverHiddenExplorerWindows", StringComparison.Ordinal),
            "StopHook must release hidden Explorer windows before disabling forced tabs.");
        Assert(disposeBody.Contains("RecoverHiddenExplorerWindows", StringComparison.Ordinal),
            "DisposeShellObjects must release hidden Explorer windows before COM objects are dropped.");
        Assert(registrationBody.Contains("RestoreMergeSourceWindowAsync(hWnd)", StringComparison.Ordinal),
            "Windows that stop being merge candidates after being hidden must be restored through the guarded restore path.");
        Assert(removeBody.Contains("Helper.RestoreHiddenExplorerWindow", StringComparison.Ordinal),
            "Removing a live shell window must restore it unless the caller intentionally closes it.");

        return Task.CompletedTask;
    }

    public static Task OrphanedTransparentExplorerWindowsAreRecoveredWithoutCache()
    {
        var helperPath = FindRepoFile("WinTab", "Helpers", "Helper.cs");
        var helper = File.ReadAllText(helperPath);
        var restoreBody = ExtractMethodBody(helper, "public static bool RestoreHiddenExplorerWindow");
        var recoveryBody = ExtractMethodBody(helper, "public static int RestoreHiddenExplorerWindows");

        Assert(restoreBody.Contains("GetLayeredWindowAttributes", StringComparison.Ordinal) &&
               restoreBody.Contains("alpha == 0", StringComparison.Ordinal) &&
               restoreBody.Contains("SetLayeredWindowAttributes(hWnd, 0, 255, WinApi.LWA_ALPHA)", StringComparison.Ordinal),
            "RestoreHiddenExplorerWindow must repair alpha=0 Explorer windows even when HiddenWindows has no cache entry.");
        Assert(recoveryBody.Contains("GetAllExplorerWindows()", StringComparison.Ordinal),
            "RestoreHiddenExplorerWindows must scan live Explorer windows, not only the in-process hidden-window cache.");
        Assert(recoveryBody.Contains("RestoreHiddenExplorerWindow", StringComparison.Ordinal),
            "RestoreHiddenExplorerWindows must route every candidate through the same restore primitive.");

        return Task.CompletedTask;
    }

    public static Task ShellInitializationSkipsDuplicateShellWindowsEntries()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var initializeBody = ExtractMethodBody(source, "private void InitializeShellObjects");

        Assert(initializeBody.Contains("_windowEntryDict.Keys.Contains(window)", StringComparison.Ordinal) &&
               initializeBody.Contains("window.GetProperty(\"seenBefore\")", StringComparison.Ordinal),
            "Startup ShellWindows enumeration can contain duplicate COM entries and must skip already-seen windows before adding them.");

        return Task.CompletedTask;
    }

    public static Task TabReuseExcludesMergeSource()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var searchBody = ExtractMethodBody(source, "private bool TrySearchForTab");
        var openBody = ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection");

        Assert(!searchBody.Contains("CreatedAt", StringComparison.Ordinal),
            "Tab reuse must not exclude freshly-created real target tabs by age.");
        Assert(searchBody.Contains("excludedTopLevelWindow", StringComparison.Ordinal),
            "Tab reuse should exclude the specific top-level source window being merged.");
        Assert(openBody.Contains("TrySearchForTab(windowToOpen.Location, windowToOpen.Handle", StringComparison.Ordinal),
            "Auto-merge must pass the source window handle into tab reuse search.");

        return Task.CompletedTask;
    }

    public static Task TabReuseDoesNotDuplicateAfterSelectionFailure()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var openBody = ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection");
        var selectBody = ExtractMethodBody(source, "private async Task<bool> SelectTabByUniqueNameVerified");

        var failureLogIndex = openBody.IndexOf("OpenTab reuse-select-failed", StringComparison.Ordinal);
        var duplicateOpenIndex = openBody.IndexOf("RequestToOpenNewTab", StringComparison.Ordinal);
        var returnTrueAfterFailureIndex = openBody.IndexOf("return true;", failureLogIndex >= 0 ? failureLogIndex : 0, StringComparison.Ordinal);
        Assert(failureLogIndex >= 0 && returnTrueAfterFailureIndex > failureLogIndex && returnTrueAfterFailureIndex < duplicateOpenIndex,
            "Once reuse finds an existing path, selection failure must not fall through to opening a duplicate tab.");
        Assert(selectBody.Contains("GetActiveTabHandle(windowHandle) == tabHandle", StringComparison.Ordinal),
            "Reuse selection should succeed immediately when the target tab is already active.");
        Assert(source.Contains("TrySelectTabByCyclingAsync", StringComparison.Ordinal) &&
               selectBody.Contains("TrySelectTabByCyclingAsync(windowHandle, tabHandle, timeoutMs)", StringComparison.Ordinal),
            "Selection failure must fall back to ExplorerTabUtility-style cycling instead of reporting reuse success without selecting the tab.");

        return Task.CompletedTask;
    }

    public static Task TabReuseForegroundsTargetBeforeSelection()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var openBody = ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection");

        var foundIndex = openBody.IndexOf("TrySearchForTab(windowToOpen.Location, windowToOpen.Handle", StringComparison.Ordinal);
        var foregroundIndex = openBody.IndexOf("WinApi.RestoreWindowToForeground(windowHandle)", foundIndex >= 0 ? foundIndex : 0, StringComparison.Ordinal);
        var selectIndex = openBody.IndexOf("SelectTabByUniqueNameVerified(windowHandle, existingTab", foundIndex >= 0 ? foundIndex : 0, StringComparison.Ordinal);

        Assert(foundIndex >= 0 && foregroundIndex > foundIndex && selectIndex > foregroundIndex,
            "External folder launches often leave a third-party app in front; reuse must foreground Explorer before selecting the matching tab.");

        return Task.CompletedTask;
    }

    public static Task MergeSourceCloseIsVerifiedBeforeHiddenTrackingIsCleared()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var registrationBody = ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");
        var closeBody = ExtractMethodBody(source, "private async Task<bool> CloseMergedSourceWindowAsync");

        var closeIndex = registrationBody.IndexOf("CloseMergedSourceWindowAsync(window, hWnd)", StringComparison.Ordinal);
        var clearIndex = registrationBody.IndexOf("RemoveMergeSourceTracking(hWnd)", closeIndex, StringComparison.Ordinal);
        Assert(closeIndex >= 0 && clearIndex > closeIndex,
            "A hidden source window must only be removed from hidden tracking after its close request has been verified.");
        Assert(closeBody.Contains("Helper.DoUntilConditionAsync", StringComparison.Ordinal) &&
               closeBody.Contains("!Helper.IsFileExplorerWindow(hWnd)", StringComparison.Ordinal) &&
               closeBody.Contains("RestoreMergeSourceWindowAsync(hWnd)", StringComparison.Ordinal),
            "Failed source-window closes must restore the Explorer window instead of leaving an alpha=0 background window.");

        return Task.CompletedTask;
    }

    public static Task FirstRunSettingsWindowFitsWithoutScrolling()
    {
        var settingsPath = FindRepoFile("WinTab", "Managers", "SettingsManager.cs");
        var xamlPath = FindRepoFile("WinTab", "UI", "Views", "MainWindow.xaml");
        var settings = File.ReadAllText(settingsPath);
        var xaml = File.ReadAllText(xamlPath);

        Assert(settings.Contains("new(1020, 720)", StringComparison.Ordinal),
            "The first-run settings size must be tall enough to show the full settings surface without a vertical scrollbar.");
        Assert(xaml.Contains("Width=\"1020\"", StringComparison.Ordinal) &&
               xaml.Contains("Height=\"720\"", StringComparison.Ordinal),
            "The XAML design size should match the persisted first-run default size.");

        return Task.CompletedTask;
    }

    public static Task SettingsResizeWritesAreDebouncedAndFlushed()
    {
        var settingsPath = FindRepoFile("WinTab", "Managers", "SettingsManager.cs");
        var mainWindowPath = FindRepoFile("WinTab", "UI", "Views", "MainWindow.xaml.cs");
        var settings = File.ReadAllText(settingsPath);
        var mainWindow = File.ReadAllText(mainWindowPath);

        Assert(settings.Contains("saveMode: SaveMode.Deferred", StringComparison.Ordinal) &&
               settings.Contains("DeferredSaveDelay", StringComparison.Ordinal) &&
               settings.Contains("_deferredSaveTimer.Change", StringComparison.Ordinal),
            "High-frequency window resize updates should debounce FormSize saves instead of writing settings.json on every SizeChanged event.");
        Assert(mainWindow.Contains("SettingsManager.SaveSettings()", StringComparison.Ordinal),
            "Application exit must flush any pending debounced settings write.");

        return Task.CompletedTask;
    }

    public static Task ExplorerHotPathsAvoidRepeatedPollingWork()
    {
        var helperPath = FindRepoFile("WinTab", "Helpers", "Helper.cs");
        var watcherPath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var helper = File.ReadAllText(helperPath);
        var watcher = File.ReadAllText(watcherPath);
        var listenBodies = ExtractMethodBody(helper, "public static Task<nint> ListenForNewExplorerWindowAsync") +
                           ExtractMethodBody(helper, "public static nint ListenForNewExplorerTab") +
                           ExtractMethodBody(helper, "public static Task<nint> ListenForNewExplorerTabAsync");
        var targetBody = ExtractMethodBody(watcher, "private nint GetMainWindowHWnd");
        var startupBody = ExtractMethodBody(watcher, "private bool IsStartupExplorerLocation");

        Assert(!listenBodies.Contains(".Except(current", StringComparison.Ordinal) &&
               listenBodies.Contains("CreateKnownHandleSet", StringComparison.Ordinal),
            "Explorer window/tab polling should reuse a known-handle set instead of allocating Except state on every poll.");
        Assert(targetBody.Contains("WinApi.FindAllWindowsEx(\"CabinetWClass\").ToArray()", StringComparison.Ordinal) &&
               targetBody.Contains("GetCachedTabCount", StringComparison.Ordinal),
            "Merge target selection should snapshot Explorer windows once and reuse tab counts across fallback passes.");
        Assert(startupBody.Contains("_startupLocationCache.TryGetValue", StringComparison.Ordinal) &&
               watcher.Contains("StartupLocationCacheLimit", StringComparison.Ordinal),
            "Startup-location checks should cache Shell PIDL equivalence results on the ShellWindows hot path without growing unbounded.");

        return Task.CompletedTask;
    }

    public static Task WindowRegisteredRestoresPreHiddenUnmergedWindows()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");

        var finallyIndex = methodBody.IndexOf("finally", StringComparison.Ordinal);
        Assert(finallyIndex >= 0, "WindowRegistered processing must have a finally block.");

        var registerIndex = methodBody.IndexOf("RegisterIndependentWindow(window, windowInfo, hWnd)", finallyIndex, StringComparison.Ordinal);
        var removedGuardIndex = methodBody.IndexOf("!removed", finallyIndex, StringComparison.Ordinal);
        var showGuardIndex = methodBody.IndexOf("showAgain", finallyIndex, StringComparison.Ordinal);

        Assert(registerIndex > finallyIndex && removedGuardIndex > finallyIndex && showGuardIndex > finallyIndex,
            "An unmerged source window must be kept as an independent visible Explorer window.");

        return Task.CompletedTask;
    }

    public static Task EarlyHideDoesNotWaitOnMergeTargetDiscovery()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var hideBody = ExtractMethodBody(source, "private bool TryHideIncomingExplorerWindow");

        Assert(!hideBody.Contains("HasMergeTargetForEarlyConceal(hWnd)", StringComparison.Ordinal),
            "Early WinEvent hide must not perform merge-target discovery before concealing a new Explorer window.");
        Assert(!hideBody.Contains("HasTrackedTopLevelWindow(hWnd)", StringComparison.Ordinal) &&
               !hideBody.Contains("HasHookedShellWindowForTopLevel(hWnd)", StringComparison.Ordinal) &&
               hideBody.Contains("_hookedTopLevelUseCounts.ContainsKey(hWnd)", StringComparison.Ordinal),
            "Early WinEvent hide must use the hooked-target cache without taking the ShellWindows dictionary lock.");

        var targetBody = ExtractMethodBody(source, "private bool IsStableMergeTargetWindow");
        Assert(targetBody.Contains("HasHookedShellWindowForTopLevel(hWnd)", StringComparison.Ordinal),
            "Merge target selection must not treat another newly-created unhooked Explorer window as a stable target.");

        return Task.CompletedTask;
    }

    public static Task WindowMergeUsesSingleRegistrationLifecycle()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var windowInfoPath = FindRepoFile("WinTab", "Models", "WindowInfo.cs");
        var source = File.ReadAllText(sourcePath);
        var windowInfo = File.ReadAllText(windowInfoPath);
        var registrationBody = ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");

        Assert(!source.Contains("SchedulePendingAutoMerge", StringComparison.Ordinal) &&
               !source.Contains("MergePendingAutoMergeWindowsAsync", StringComparison.Ordinal) &&
               !source.Contains("GetPendingAutoMergeCandidates", StringComparison.Ordinal),
            "Hidden Explorer windows must not be owned by a background pending merge queue.");
        Assert(!windowInfo.Contains("CanAutoMerge", StringComparison.Ordinal),
            "WindowInfo must not carry a long-lived pending merge state.");
        Assert(source.Contains("private bool HasUntrackedShellWindows()", StringComparison.Ordinal),
            "Registration processing must close the race where a ShellWindows event arrives during an active drain.");

        var openIndex = registrationBody.IndexOf("OpenTabNavigateWithSelection(record, targetWindow)", StringComparison.Ordinal);
        var quitIndex = registrationBody.IndexOf("CloseMergedSourceWindowAsync(window, hWnd)", openIndex, StringComparison.Ordinal);
        var trackingIndex = registrationBody.IndexOf("RemoveMergeSourceTracking(hWnd)", quitIndex, StringComparison.Ordinal);
        var removeIndex = registrationBody.IndexOf("RemoveWindowAndUnhookEvents(window, windowInfo, restoreHiddenWindow: false)", trackingIndex, StringComparison.Ordinal);
        Assert(openIndex >= 0 && quitIndex > openIndex && trackingIndex > quitIndex && removeIndex > trackingIndex,
            "The source Explorer window must be closed and verified after the target tab opens, before hidden tracking is cleared.");

        var hideIndex = registrationBody.IndexOf("HideMergeSourceWindow(hWnd)", StringComparison.Ordinal);
        var tabHandleIndex = registrationBody.IndexOf("GetTabHandle(window)", StringComparison.Ordinal);
        var tabCountIndex = registrationBody.IndexOf("WaitForExplorerTabCount(hWnd)", StringComparison.Ordinal);
        var resolveIndex = registrationBody.IndexOf("ResolveInitialLocation(window, hWnd)", StringComparison.Ordinal);
        Assert(hideIndex >= 0 && hideIndex < tabHandleIndex && tabCountIndex > hideIndex && resolveIndex > hideIndex,
            "A mergeable source window must be hidden before slower tab-handle, tab-count, and location waits can let Explorer paint it.");
        Assert(registrationBody.Contains("RemoveMergeSourceTracking(hWnd)", StringComparison.Ordinal),
            "A successfully merged source window must be removed from hidden-source tracking after the close is verified.");

        return Task.CompletedTask;
    }

    public static Task UpdateCheckPrefersCurrentArchitectureInstaller()
    {
        var sourcePath = FindRepoFile("WinTab", "Managers", "UpdateManager.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "private static string? FindMatchingUpdateAssetUrl") +
                         ExtractMethodBody(source, "private static string? GetInstallerArchitectureSuffix");

        Assert(methodBody.Contains("RuntimeInformation.ProcessArchitecture", StringComparison.Ordinal) &&
               methodBody.Contains("Architecture.X64", StringComparison.Ordinal) &&
               methodBody.Contains("Architecture.X86", StringComparison.Ordinal) &&
               methodBody.Contains("Architecture.Arm64", StringComparison.Ordinal),
            "Update downloads must match the current process architecture instead of using the first installer asset.");
        Assert(methodBody.Contains("_x64_Setup.exe", StringComparison.Ordinal) &&
               methodBody.Contains("_x86_Setup.exe", StringComparison.Ordinal) &&
               methodBody.Contains("_arm64_Setup.exe", StringComparison.Ordinal),
            "Update downloads must select the release asset suffix for each packaged architecture.");

        return Task.CompletedTask;
    }

    private static string FindRepoFile(params string[] relativeParts)
    {
        foreach (var start in new[] { Environment.CurrentDirectory, AppContext.BaseDirectory })
        {
            var current = new DirectoryInfo(start);
            while (current != null)
            {
                var candidate = Path.Combine(new[] { current.FullName }.Concat(relativeParts).ToArray());
                if (File.Exists(candidate))
                    return candidate;

                current = current.Parent;
            }
        }

        throw new FileNotFoundException("Could not locate repo file.", Path.Combine(relativeParts));
    }

    private static string ExtractMethodBody(string source, string signature)
    {
        var signatureIndex = source.IndexOf(signature, StringComparison.Ordinal);
        if (signatureIndex < 0)
            throw new InvalidOperationException($"Could not find method signature '{signature}'.");

        var openBraceIndex = source.IndexOf('{', signatureIndex);
        if (openBraceIndex < 0)
            throw new InvalidOperationException($"Could not find method body for '{signature}'.");

        var depth = 0;
        for (var i = openBraceIndex; i < source.Length; i++)
        {
            if (source[i] == '{')
                depth++;
            else if (source[i] == '}')
                depth--;

            if (depth == 0)
                return source.Substring(openBraceIndex, i - openBraceIndex + 1);
        }

        throw new InvalidOperationException($"Could not parse method body for '{signature}'.");
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }
}

internal static class ExplorerStressTest
{
    private static readonly Guid ShellBrowserGuid = typeof(IShellBrowser).GUID;

    public static async Task<int> RunAsync(string[] args)
    {
        var appPath = GetOption(args, "--app");
        if (string.IsNullOrWhiteSpace(appPath) || !File.Exists(appPath))
            throw new InvalidOperationException("Pass --app with the WinTab.exe path to stress-test.");

        var roundCount = int.TryParse(GetOption(args, "--rounds"), out var parsedRounds) ? parsedRounds : 7;
        var intervalMs = int.TryParse(GetOption(args, "--interval"), out var parsedInterval) ? parsedInterval : 75;
        var startupDelayMs = int.TryParse(GetOption(args, "--startup-delay"), out var parsedStartupDelay) ? parsedStartupDelay : 3_000;
        var failOnVisibleNewWindow = HasOption(args, "--fail-on-visible-new-window");
        var failOnTransientDefaultTab = HasOption(args, "--fail-on-transient-default-tab");
        var failOnExplorerRestart = HasOption(args, "--fail-on-explorer-restart");
        var stressRoot = Path.Combine(Path.GetTempPath(), "WinTabExplorerStress");
        CloseTestShellWindows(stressRoot);

        var root = Path.Combine(stressRoot, DateTime.Now.ToString("yyyyMMddHHmmssfff"));
        var debugLog = Path.Combine(root, "wintab-debug.log");
        var baseline = CreateBaseline(root);
        var targets = CreateTargets(root, roundCount);
        StartExplorer(baseline);
        await WaitForFolderWindowAsync(baseline);

        var before = GetShellWindows();
        var beforeDefaultCount = before.Count(IsDefaultLocation);
        var beforeBaselineDefaultCount = before.Count(window => IsDefaultLocation(window));
        var baselineTopLevelWindows = Helper.GetAllExplorerWindows()
            .Concat(before.Select(window => (nint)window.Hwnd))
            .ToHashSet();
        VisibleExplorerWindowMonitor? visibleWindowMonitor = null;
        TransientDefaultTabMonitor? transientDefaultTabMonitor = null;
        ExplorerProcessMonitor? explorerProcessMonitor = null;

        using var app = StartWinTab(appPath, debugLog);
        try
        {
            await Task.Delay(startupDelayMs);
            if (failOnVisibleNewWindow)
                visibleWindowMonitor = VisibleExplorerWindowMonitor.Start(baselineTopLevelWindows);
            if (failOnTransientDefaultTab)
                transientDefaultTabMonitor = TransientDefaultTabMonitor.Start(baselineTopLevelWindows, beforeBaselineDefaultCount);
            if (failOnExplorerRestart)
                explorerProcessMonitor = ExplorerProcessMonitor.Start(baselineTopLevelWindows);

            foreach (var target in targets)
            {
                StartExplorer(target);
                await Task.Delay(intervalMs);
            }

            var finalWindows = await WaitForTargetsAsync(targets, beforeDefaultCount, baselineTopLevelWindows);
            var missingTargets = targets.Where(target => finalWindows.All(window => !IsSameFolder(window, target))).ToArray();
            var defaultCount = finalWindows.Count(IsDefaultLocation);
            var unmergedTargetWindows = GetUnexpectedTargetTopLevelWindows(finalWindows, targets, baselineTopLevelWindows).ToArray();
            var visibleNewWindows = visibleWindowMonitor?.Sightings.ToArray() ?? [];
            var transientDefaultTabs = transientDefaultTabMonitor?.Sightings.ToArray() ?? [];
            var explorerRestarts = explorerProcessMonitor?.ExitedProcessIds.ToArray() ?? [];

            if (missingTargets.Length > 0 ||
                defaultCount > beforeDefaultCount ||
                unmergedTargetWindows.Length > 0 ||
                visibleNewWindows.Length > 0 ||
                transientDefaultTabs.Length > 0 ||
                explorerRestarts.Length > 0)
            {
                Console.Error.WriteLine("Explorer stress test failed.");
                Console.Error.WriteLine($"Missing targets: {string.Join(", ", missingTargets)}");
                Console.Error.WriteLine($"Default-location windows before={beforeDefaultCount}, after={defaultCount}");
                if (explorerRestarts.Length > 0)
                    Console.Error.WriteLine($"Explorer process exited during stress: {string.Join(", ", explorerRestarts)}");
                if (unmergedTargetWindows.Length > 0)
                {
                    Console.Error.WriteLine("Targets still live in new top-level Explorer windows:");
                    DumpShellWindows(unmergedTargetWindows);
                }
                if (visibleNewWindows.Length > 0)
                {
                    Console.Error.WriteLine("New Explorer windows became visible on screen before merging:");
                    foreach (var sighting in visibleNewWindows)
                        Console.Error.WriteLine($"HWND={sighting.Hwnd} Rect={sighting.Left},{sighting.Top},{sighting.Right},{sighting.Bottom} First={sighting.FirstObservedAt:HH:mm:ss.fff} Last={sighting.LastObservedAt:HH:mm:ss.fff} DurationMs={sighting.VisibleDuration.TotalMilliseconds:F1} Samples={sighting.Observations}");
                }
                if (transientDefaultTabs.Length > 0)
                {
                    Console.Error.WriteLine("Existing Explorer window showed transient active default-location tabs:");
                    foreach (var sighting in transientDefaultTabs)
                        Console.Error.WriteLine($"HWND={sighting.Hwnd} Count={sighting.DefaultCount} First={sighting.FirstObservedAt:HH:mm:ss.fff} Last={sighting.LastObservedAt:HH:mm:ss.fff} DurationMs={sighting.VisibleDuration.TotalMilliseconds:F1} Samples={sighting.Observations}");
                }
                DumpShellWindows(finalWindows);
                DumpDebugLog(debugLog);
                return 1;
            }

            Console.WriteLine($"PASS Explorer stress: {targets.Length} rapid folder opens resolved to target folders.");
            Console.WriteLine($"PASS Default-location windows did not increase: before={beforeDefaultCount}, after={defaultCount}.");
            if (baselineTopLevelWindows.Count > 0)
                Console.WriteLine("PASS Target folders ended in existing Explorer top-level windows.");
            if (failOnVisibleNewWindow)
                Console.WriteLine("PASS New Explorer top-level windows were concealed before a visible merge flash threshold.");
            if (failOnTransientDefaultTab)
                Console.WriteLine("PASS Existing Explorer windows did not show transient active default-location tabs.");
            if (failOnExplorerRestart)
                Console.WriteLine("PASS Existing Explorer process stayed alive during stress.");
            return 0;
        }
        finally
        {
            if (visibleWindowMonitor != null)
                await visibleWindowMonitor.StopAsync();
            if (transientDefaultTabMonitor != null)
                await transientDefaultTabMonitor.StopAsync();
            if (explorerProcessMonitor != null)
                await explorerProcessMonitor.StopAsync();
            CloseTestShellWindows(root);
            TryKill(app);
            TryDelete(root);
        }
    }

    public static async Task<int> RunActivationAsync(string[] args)
    {
        var tabCount = int.TryParse(GetOption(args, "--tabs"), out var parsedTabs) ? parsedTabs : 8;
        var maxActivationMs = int.TryParse(GetOption(args, "--max-activation-ms"), out var parsedMaxActivationMs) ? parsedMaxActivationMs : 700;
        var stressRoot = Path.Combine(Path.GetTempPath(), "WinTabExplorerStress");
        CloseTestShellWindows(stressRoot);

        var root = Path.Combine(stressRoot, DateTime.Now.ToString("yyyyMMddHHmmssfff"));
        var baseline = CreateBaseline(root);
        StartExplorer(baseline);
        await WaitForFolderWindowAsync(baseline);
        nint controlledWindow = 0;

        try
        {
            var baselineWindow = GetShellWindows().FirstOrDefault(window => IsSameFolder(window, baseline));
            if (baselineWindow == null)
            {
                Console.Error.WriteLine("Explorer activation stress test failed: no baseline window found.");
                return 1;
            }

            var mainHwnd = (nint)baselineWindow.Hwnd;
            controlledWindow = mainHwnd;
            var targetTab = GetActiveTabHandle(mainHwnd);
            if (targetTab == 0)
            {
                Console.Error.WriteLine("Explorer activation stress test failed: could not resolve the baseline tab.");
                DumpShellWindows(GetShellWindows());
                return 1;
            }

            if (!await CreateControlledBlankTabsAsync(mainHwnd, Math.Max(0, tabCount - 1)))
            {
                Console.Error.WriteLine("Explorer activation stress test failed: could not create controlled Explorer tabs.");
                DumpShellWindows(GetShellWindows());
                return 1;
            }

            var tabs = Helper.GetAllExplorerTabs(mainHwnd).ToArray();
            if (tabs.Length < Math.Min(4, tabCount))
            {
                Console.Error.WriteLine($"Explorer activation stress test failed: expected multiple tabs, found {tabs.Length}.");
                DumpShellWindows(GetShellWindows());
                return 1;
            }

            using var watcher = new ExplorerWatcher();
            await Helper.DoUntilConditionAsync(() => watcher.IsShellReady, ready => ready, 3_000, 50);
            await Task.Delay(1_500);

            watcher.SelectLastTab(mainHwnd);
            var startTab = await WaitForAnyActiveTabAsync(mainHwnd, 2_000);
            if (startTab == 0 || startTab == targetTab)
            {
                Console.Error.WriteLine("Explorer activation stress test failed: could not select the start tab.");
                return 1;
            }

            var monitor = ActiveTabSequenceMonitor.Start(mainHwnd);
            var stopwatch = Stopwatch.StartNew();
            var selectionTask = watcher.SelectTabByHandle(mainHwnd, targetTab);
            var selected = await Task.WhenAny(selectionTask, Task.Delay(maxActivationMs)) == selectionTask;
            if (selected)
                await selectionTask;
            stopwatch.Stop();
            await Task.Delay(250);
            await monitor.StopAsync();

            var finalActiveTab = GetActiveTabHandle(mainHwnd);
            var intermediateTabs = monitor.ObservedTabs
                .Where(tab => tab != 0 && tab != startTab && tab != targetTab)
                .Distinct()
                .ToArray();
            if (!selected ||
                finalActiveTab != targetTab ||
                intermediateTabs.Length > 0 ||
                stopwatch.ElapsedMilliseconds > maxActivationMs)
            {
                Console.Error.WriteLine("Explorer activation stress test failed.");
                Console.Error.WriteLine($"Selected={selected} FinalActive={finalActiveTab} Target={targetTab} ElapsedMs={stopwatch.ElapsedMilliseconds} MaxMs={maxActivationMs}");
                Console.Error.WriteLine($"Intermediate active tabs: {string.Join(", ", intermediateTabs)}");
                Console.Error.WriteLine($"Observed active tabs: {string.Join(", ", monitor.ObservedTabs)}");
                return 1;
            }

            Console.WriteLine($"PASS Explorer activation: selected target tab directly in {stopwatch.ElapsedMilliseconds} ms without intermediate active tabs.");
            return 0;
        }
        finally
        {
            CloseShellWindowByHwnd(controlledWindow);
            CloseTestShellWindows(root);
            TryDelete(root);
        }
    }

    public static async Task<int> RunRecoveryAsync(string[] args)
    {
        var appPath = GetOption(args, "--app");
        if (string.IsNullOrWhiteSpace(appPath) || !File.Exists(appPath))
            throw new InvalidOperationException("Pass --app with the WinTab.exe path to recovery-test.");

        var startupDelayMs = int.TryParse(GetOption(args, "--startup-delay"), out var parsedStartupDelay) ? parsedStartupDelay : 2_500;
        var stressRoot = Path.Combine(Path.GetTempPath(), "WinTabExplorerStress");
        KillExistingWinTabProcesses();
        CloseTestShellWindows(stressRoot);

        var root = Path.Combine(stressRoot, DateTime.Now.ToString("yyyyMMddHHmmssfff"));
        var debugLog = Path.Combine(root, "wintab-recovery-debug.log");
        var baseline = CreateBaseline(root);
        StartExplorerNewWindow(baseline);
        await WaitForFolderWindowAsync(baseline);

        nint concealedWindow = 0;
        Process? app = null;
        try
        {
            var baselineWindow = GetShellWindows().FirstOrDefault(window => IsSameFolder(window, baseline));
            if (baselineWindow == null)
            {
                Console.Error.WriteLine("Explorer recovery test failed: no baseline Explorer window found.");
                return 1;
            }

            concealedWindow = (nint)baselineWindow.Hwnd;
            ConcealExplorerWindowAsOrphan(concealedWindow);
            await Helper.DoUntilConditionAsync(
                () => IsFullyTransparent(concealedWindow),
                transparent => transparent,
                1_500,
                50);

            if (!IsFullyTransparent(concealedWindow))
            {
                Console.Error.WriteLine("Explorer recovery test failed: could not reproduce alpha=0 Explorer window.");
                DumpShellWindows(GetShellWindows());
                return 1;
            }

            app = StartWinTab(appPath, debugLog);
            await Task.Delay(startupDelayMs);

            var restored = await Helper.DoUntilConditionAsync(
                () => IsExplorerWindowRestored(concealedWindow),
                isRestored => isRestored,
                6_000,
                50);

            if (!restored)
            {
                Console.Error.WriteLine("Explorer recovery test failed: WinTab did not restore an orphaned alpha=0 Explorer window.");
                DumpWindowVisibility(concealedWindow);
                DumpShellWindows(GetShellWindows());
                DumpDebugLog(debugLog);
                return 1;
            }

            Console.WriteLine("PASS Explorer recovery: orphaned alpha=0 Explorer window was restored by WinTab startup.");
            return 0;
        }
        finally
        {
            TryRestoreExplorerWindow(concealedWindow);
            if (app != null)
                TryKill(app);
            CloseTestShellWindows(root);
            TryDelete(root);
        }
    }

    public static async Task<int> RunReuseAsync(string[] args)
    {
        var appPath = GetOption(args, "--app");
        if (string.IsNullOrWhiteSpace(appPath) || !File.Exists(appPath))
            throw new InvalidOperationException("Pass --app with the WinTab.exe path to reuse-test.");

        var targetOverride = GetOption(args, "--target");
        var startupDelayMs = int.TryParse(GetOption(args, "--startup-delay"), out var parsedStartupDelay) ? parsedStartupDelay : 2_500;
        var repeatDelayMs = int.TryParse(GetOption(args, "--repeat-delay"), out var parsedRepeatDelay) ? parsedRepeatDelay : 250;
        var externalShellOpen = HasOption(args, "--external-shell");
        var stressRoot = Path.Combine(Path.GetTempPath(), "WinTabExplorerStress");
        KillExistingWinTabProcesses();
        CloseTestShellWindows(stressRoot);

        var root = Path.Combine(stressRoot, DateTime.Now.ToString("yyyyMMddHHmmssfff"));
        var debugLog = Path.Combine(root, "wintab-reuse-debug.log");
        var baseline = CreateBaseline(root);
        var target = string.IsNullOrWhiteSpace(targetOverride)
            ? Path.Combine(root, "新建文件夹")
            : targetOverride;

        if (!Directory.Exists(target))
            Directory.CreateDirectory(target);
        if (!string.IsNullOrWhiteSpace(targetOverride))
            CloseShellWindowsByFolder(target);

        StartExplorer(baseline);
        await WaitForFolderWindowAsync(baseline);
        var before = GetShellWindows();
        var beforeDefaultCount = before.Count(IsDefaultLocation);
        var baselineTopLevelWindows = Helper.GetAllExplorerWindows()
            .Concat(before.Select(window => (nint)window.Hwnd))
            .ToHashSet();

        using var app = StartWinTab(appPath, debugLog);
        try
        {
            await Task.Delay(startupDelayMs);

            StartFolder(target, externalShellOpen);
            var firstTargetWindow = await WaitForTargetsAsync([target], beforeDefaultCount, baselineTopLevelWindows);
            var unmergedFirstTargetWindows = GetUnexpectedTargetTopLevelWindows(firstTargetWindow, [target], baselineTopLevelWindows);
            if (firstTargetWindow.Count(window => IsSameFolder(window, target)) != 1 ||
                unmergedFirstTargetWindows.Length > 0)
            {
                Console.Error.WriteLine("Explorer reuse test failed: first open did not merge into the existing Explorer tab set.");
                DumpShellWindows(firstTargetWindow);
                DumpDebugLog(debugLog);
                return 1;
            }

            var firstTargetTab = firstTargetWindow.First(window => IsSameFolder(window, target));
            var targetWindowHandle = (nint)firstTargetTab.Hwnd;
            var targetTabHandle = GetActiveTabHandle(targetWindowHandle);
            if (targetTabHandle == 0 ||
                !await SelectAnyOtherTabAsync(targetWindowHandle, targetTabHandle, 2_000))
            {
                Console.Error.WriteLine("Explorer reuse test failed: could not move away from the first target tab before reopening it.");
                Console.Error.WriteLine($"TargetWindow={targetWindowHandle} TargetTab={targetTabHandle}");
                DumpShellWindows(GetShellWindows());
                DumpDebugLog(debugLog);
                return 1;
            }

            await Task.Delay(repeatDelayMs);
            StartFolder(target, externalShellOpen);
            var finalWindows = await WaitForTargetCountAsync(target, expectedCount: 1, timeoutMs: 12_000);
            var targetTabs = finalWindows.Where(window => IsSameFolder(window, target)).ToArray();

            if (targetTabs.Length != 1)
            {
                Console.Error.WriteLine("Explorer reuse test failed: reopening the same folder created duplicate target tabs.");
                Console.Error.WriteLine($"Target='{target}' Count={targetTabs.Length}");
                DumpShellWindows(finalWindows);
                DumpDebugLog(debugLog);
                return 1;
            }

            var activeTarget = await WaitForActiveTabAsync((nint)targetTabs[0].Hwnd, targetTabHandle, 2_000);
            if (activeTarget != targetTabHandle)
            {
                Console.Error.WriteLine("Explorer reuse test failed: reopening the same folder did not activate the existing target tab.");
                Console.Error.WriteLine($"ExpectedActiveTab={targetTabHandle} ActualActiveTab={activeTarget}");
                DumpShellWindows(finalWindows);
                DumpDebugLog(debugLog);
                return 1;
            }

            var mode = externalShellOpen ? "external Shell open" : "explorer.exe";
            Console.WriteLine($"PASS Explorer reuse ({mode}): reopening '{target}' activated the existing tab without creating a duplicate.");
            return 0;
        }
        finally
        {
            TryKill(app);
            if (string.IsNullOrWhiteSpace(targetOverride))
                CloseTestShellWindows(root);
            else
                CloseShellWindowsByFolder(target);
            TryDelete(root);
        }
    }

    public static async Task<int> RunDefaultLocationAsync(string[] args)
    {
        var appPath = GetOption(args, "--app");
        if (string.IsNullOrWhiteSpace(appPath) || !File.Exists(appPath))
            throw new InvalidOperationException("Pass --app with the WinTab.exe path to default-location-test.");

        var startupDelayMs = int.TryParse(GetOption(args, "--startup-delay"), out var parsedStartupDelay) ? parsedStartupDelay : 2_500;
        var observeMs = int.TryParse(GetOption(args, "--observe-ms"), out var parsedObserveMs) ? parsedObserveMs : 5_500;
        var stressRoot = Path.Combine(Path.GetTempPath(), "WinTabExplorerStress");
        KillExistingWinTabProcesses();
        CloseTestShellWindows(stressRoot);

        var root = Path.Combine(stressRoot, DateTime.Now.ToString("yyyyMMddHHmmssfff"));
        var debugLog = Path.Combine(root, "wintab-default-location-debug.log");
        var baseline = CreateBaseline(root);
        StartExplorer(baseline);
        await WaitForFolderWindowAsync(baseline);

        using var app = StartWinTab(appPath, debugLog);
        var defaultWindowsBefore = GetShellWindows()
            .Where(IsDefaultLocation)
            .Select(window => (nint)window.Hwnd)
            .ToHashSet();
        var defaultWindowsOpenedByTest = new HashSet<nint>();

        try
        {
            await Task.Delay(startupDelayMs);

            StartExplorerDefault();
            await Task.Delay(observeMs);

            var windows = GetShellWindows();
            var defaultWindows = windows.Where(IsDefaultLocation).ToArray();
            foreach (var window in defaultWindows)
            {
                var hWnd = (nint)window.Hwnd;
                if (!defaultWindowsBefore.Contains(hWnd))
                    defaultWindowsOpenedByTest.Add(hWnd);
            }

            var transparentDefaultWindows = defaultWindows
                .Where(window => IsFullyTransparent((nint)window.Hwnd))
                .ToArray();

            if (transparentDefaultWindows.Length > 0)
            {
                Console.Error.WriteLine("Explorer default-location test failed.");
                Console.Error.WriteLine($"Default windows found={defaultWindows.Length}, transparent defaults={transparentDefaultWindows.Length}.");
                DumpShellWindows(windows);
                foreach (var window in transparentDefaultWindows)
                    DumpWindowVisibility((nint)window.Hwnd);
                DumpDebugLog(debugLog);
                return 1;
            }

            Console.WriteLine("PASS Explorer default-location: opening This PC did not leave a transparent hidden Explorer window.");
            return 0;
        }
        finally
        {
            foreach (var hWnd in defaultWindowsOpenedByTest)
            {
                TryRestoreExplorerWindow(hWnd);
                CloseShellWindowByHwnd(hWnd);
            }

            TryKill(app);
            CloseTestShellWindows(root);
            TryDelete(root);
        }
    }

    public static async Task<int> RunUserDefaultAsync(string[] args)
    {
        var appPath = GetOption(args, "--app");
        if (string.IsNullOrWhiteSpace(appPath) || !File.Exists(appPath))
            throw new InvalidOperationException("Pass --app with the WinTab.exe path to user-default-test.");

        var startupDelayMs = int.TryParse(GetOption(args, "--startup-delay"), out var parsedStartupDelay) ? parsedStartupDelay : 2_500;
        var observeMs = int.TryParse(GetOption(args, "--observe-ms"), out var parsedObserveMs) ? parsedObserveMs : 3_500;
        var maxHideMs = int.TryParse(GetOption(args, "--max-hide-ms"), out var parsedMaxHideMs) ? parsedMaxHideMs : 1_000;
        var stressRoot = Path.Combine(Path.GetTempPath(), "WinTabExplorerStress");
        KillExistingWinTabProcesses();
        CloseTestShellWindows(stressRoot);

        var root = Path.Combine(stressRoot, DateTime.Now.ToString("yyyyMMddHHmmssfff"));
        var debugLog = Path.Combine(root, "wintab-user-default-debug.log");
        var baseline = CreateBaseline(root);
        StartExplorer(baseline);
        await WaitForFolderWindowAsync(baseline);

        using var app = StartWinTab(appPath, debugLog);
        var baselineDefaultHwnds = GetShellWindows()
            .Where(IsDefaultLocation)
            .Select(window => (nint)window.Hwnd)
            .ToHashSet();
        var baselineTopLevelHwnds = GetShellWindows()
            .Select(window => (nint)window.Hwnd)
            .ToHashSet();
        var openedByTest = new HashSet<nint>();

        try
        {
            await Task.Delay(startupDelayMs);

            var monitor = HiddenDefaultWindowMonitor.Start(baselineTopLevelHwnds);
            StartExplorerDefault();
            await Task.Delay(observeMs);
            await monitor.StopAsync();

            var windows = GetShellWindows();
            var defaultWindows = windows.Where(IsDefaultLocation).ToArray();
            var newDefaultTopLevelWindows = defaultWindows
                .Where(window => !baselineTopLevelHwnds.Contains((nint)window.Hwnd))
                .ToArray();
            foreach (var hwnd in newDefaultTopLevelWindows.Select(window => (nint)window.Hwnd))
            {
                openedByTest.Add(hwnd);
            }

            var transparentDefaultWindows = defaultWindows
                .Where(window => IsFullyTransparent((nint)window.Hwnd))
                .ToArray();
            var sightings = monitor.Sightings.ToArray();
            var visibleIntermediates = sightings.Where(sighting => sighting.FirstVisibleAt.HasValue).ToArray();
            var offending = sightings.Where(sighting => sighting.HiddenMs > maxHideMs).ToArray();
            if (newDefaultTopLevelWindows.Length > 0 ||
                transparentDefaultWindows.Length > 0 ||
                visibleIntermediates.Length > 0 ||
                offending.Length > 0)
            {
                Console.Error.WriteLine("Explorer user-default test failed: This PC showed a visible intermediate window, did not merge, or left a hidden residual.");
                Console.Error.WriteLine($"New default top-level windows={newDefaultTopLevelWindows.Length}");
                Console.Error.WriteLine($"Visible intermediate windows={visibleIntermediates.Length}");
                Console.Error.WriteLine($"Threshold={maxHideMs}ms");
                foreach (var sighting in sightings)
                {
                    Console.Error.WriteLine($"HWND={sighting.Hwnd} HiddenMs={sighting.HiddenMs:F1} FirstSeen={sighting.FirstSeen:HH:mm:ss.fff} HiddenAt={(sighting.FirstHiddenAt?.ToString("HH:mm:ss.fff") ?? "none")} VisibleAt={(sighting.FirstVisibleAt?.ToString("HH:mm:ss.fff") ?? "none")} LastHiddenAt={(sighting.LastHiddenAt?.ToString("HH:mm:ss.fff") ?? "none")}");
                }
                DumpShellWindows(windows);
                foreach (var window in transparentDefaultWindows)
                    DumpWindowVisibility((nint)window.Hwnd);
                DumpDebugLog(debugLog);
                return 1;
            }

            var maxObserved = sightings.Length == 0 ? 0 : sightings.Max(sighting => sighting.HiddenMs);
            Console.WriteLine($"PASS Explorer user-default: user-opened This PC merged into an existing target without visible intermediate UI or hidden residuals (max observed={maxObserved:F0}ms).");
            return 0;
        }
        finally
        {
            foreach (var hWnd in openedByTest)
            {
                TryRestoreExplorerWindow(hWnd);
                CloseShellWindowByHwnd(hWnd);
            }

            TryKill(app);
            CloseTestShellWindows(root);
            TryDelete(root);
        }
    }

    public static async Task<int> RunMixedDefaultFolderAsync(string[] args)
    {
        var appPath = GetOption(args, "--app");
        if (string.IsNullOrWhiteSpace(appPath) || !File.Exists(appPath))
            throw new InvalidOperationException("Pass --app with the WinTab.exe path to mixed-default-folder-test.");

        var roundCount = int.TryParse(GetOption(args, "--rounds"), out var parsedRounds) ? parsedRounds : 8;
        var intervalMs = int.TryParse(GetOption(args, "--interval"), out var parsedInterval) ? parsedInterval : 35;
        var startupDelayMs = int.TryParse(GetOption(args, "--startup-delay"), out var parsedStartupDelay) ? parsedStartupDelay : 2_500;
        var defaultLeadMs = int.TryParse(GetOption(args, "--default-lead-ms"), out var parsedDefaultLead) ? parsedDefaultLead : 120;
        var maxResolveMs = int.TryParse(GetOption(args, "--max-resolve-ms"), out var parsedMaxResolve) ? parsedMaxResolve : 0;
        var stressRoot = Path.Combine(Path.GetTempPath(), "WinTabExplorerStress");
        KillExistingWinTabProcesses();
        CloseTestShellWindows(stressRoot);

        var root = Path.Combine(stressRoot, DateTime.Now.ToString("yyyyMMddHHmmssfff"));
        var debugLog = Path.Combine(root, "wintab-mixed-debug.log");
        var baseline = CreateBaseline(root);
        var targets = CreateTargets(root, roundCount);
        StartExplorer(baseline);
        await WaitForFolderWindowAsync(baseline);

        var before = GetShellWindows();
        var beforeDefaultCount = before.Count(IsDefaultLocation);
        var baselineTopLevelWindows = Helper.GetAllExplorerWindows()
            .Concat(before.Select(window => (nint)window.Hwnd))
            .ToHashSet();

        using var app = StartWinTab(appPath, debugLog);
        try
        {
            await Task.Delay(startupDelayMs);

            var stopwatch = Stopwatch.StartNew();
            StartExplorerDefault();
            await Task.Delay(defaultLeadMs);
            foreach (var target in targets)
            {
                StartExplorer(target);
                await Task.Delay(intervalMs);
            }

            var finalWindows = await WaitForTargetsAsync(targets, beforeDefaultCount + 1, baselineTopLevelWindows);
            stopwatch.Stop();

            var missingTargets = targets.Where(target => finalWindows.All(window => !IsSameFolder(window, target))).ToArray();
            var defaultCount = finalWindows.Count(IsDefaultLocation);
            var transparentDefaultWindows = finalWindows
                .Where(IsDefaultLocation)
                .Where(window => IsFullyTransparent((nint)window.Hwnd))
                .ToArray();
            var unmergedTargetWindows = GetUnexpectedTargetTopLevelWindows(finalWindows, targets, baselineTopLevelWindows).ToArray();
            var tooSlow = maxResolveMs > 0 && stopwatch.ElapsedMilliseconds > maxResolveMs;

            if (missingTargets.Length > 0 ||
                defaultCount > beforeDefaultCount + 1 ||
                transparentDefaultWindows.Length > 0 ||
                unmergedTargetWindows.Length > 0 ||
                tooSlow)
            {
                Console.Error.WriteLine("Explorer mixed default/folder stress test failed.");
                Console.Error.WriteLine($"Missing targets: {string.Join(", ", missingTargets)}");
                Console.Error.WriteLine($"Default-location windows before={beforeDefaultCount}, after={defaultCount}, allowed={beforeDefaultCount + 1}.");
                Console.Error.WriteLine($"ElapsedMs={stopwatch.ElapsedMilliseconds}, MaxResolveMs={maxResolveMs}");
                if (unmergedTargetWindows.Length > 0)
                {
                    Console.Error.WriteLine("Targets still live in new top-level Explorer windows:");
                    DumpShellWindows(unmergedTargetWindows);
                }
                foreach (var window in transparentDefaultWindows)
                    DumpWindowVisibility((nint)window.Hwnd);
                DumpShellWindows(finalWindows);
                DumpDebugLog(debugLog);
                return 1;
            }

            Console.WriteLine($"PASS Explorer mixed default/folder: active This PC plus {targets.Length} rapid folders produced no hidden default residuals or extra default pollution in {stopwatch.ElapsedMilliseconds} ms.");
            return 0;
        }
        finally
        {
            TryKill(app);
            CloseTestShellWindows(root);
            TryDelete(root);
        }
    }

    private static string[] CreateTargets(string root, int count)
    {
        Directory.CreateDirectory(root);

        var targets = new string[count];
        for (var i = 0; i < count; i++)
        {
            var path = Path.Combine(root, $"Target-{i + 1:00}");
            Directory.CreateDirectory(path);
            File.WriteAllText(Path.Combine(path, "marker.txt"), path);
            targets[i] = path;
        }

        return targets;
    }
    private static string CreateBaseline(string root)
    {
        var path = Path.Combine(root, "Baseline");
        Directory.CreateDirectory(path);
        File.WriteAllText(Path.Combine(path, "marker.txt"), path);
        return path;
    }

    private static Process StartWinTab(string appPath, string debugLog)
    {
        KillExistingWinTabProcesses();

        var startInfo = new ProcessStartInfo(appPath, "--background")
        {
            UseShellExecute = false,
            WorkingDirectory = Path.GetDirectoryName(appPath) ?? Environment.CurrentDirectory
        };
        startInfo.Environment["WINTAB_DEBUG_LOG"] = debugLog;

        var process = Process.Start(startInfo);
        if (process == null)
            throw new InvalidOperationException("Could not start WinTab.");

        return process;
    }

    private static void StartExplorer(string target)
    {
        Process.Start(new ProcessStartInfo("explorer.exe", $"\"{target}\"") { UseShellExecute = false });
    }
    private static void StartExplorerNewWindow(string target)
    {
        Process.Start(new ProcessStartInfo("explorer.exe", $"/n,\"{target}\"") { UseShellExecute = false });
    }
    private static void StartExplorerDefault()
    {
        Process.Start(new ProcessStartInfo("explorer.exe") { UseShellExecute = false });
    }
    private static void StartFolder(string target, bool externalShellOpen)
    {
        if (externalShellOpen)
        {
            Process.Start(new ProcessStartInfo(target) { UseShellExecute = true, Verb = "open" });
            return;
        }

        StartExplorer(target);
    }
    private static async Task WaitForFolderWindowAsync(string folder)
    {
        var timeoutAt = Environment.TickCount64 + 8_000;
        while (Environment.TickCount64 < timeoutAt)
        {
            if (GetShellWindows().Any(window => IsSameFolder(window, folder)))
                return;

            await Task.Delay(100);
        }

        throw new InvalidOperationException($"Could not open baseline Explorer window for '{folder}'.");
    }

    private static async Task<bool> CreateControlledBlankTabsAsync(nint mainHwnd, int additionalTabCount)
    {
        var knownTabs = Helper.GetAllExplorerTabs(mainHwnd).ToArray();
        for (var i = 0; i < additionalTabCount; i++)
        {
            var activeTab = GetActiveTabHandle(mainHwnd);
            if (activeTab == 0)
                return false;

            WinApi.PostMessage(activeTab, WinApi.WM_COMMAND, 0xA21B, 0);
            var newTab = await Helper.ListenForNewExplorerTabAsync(mainHwnd, knownTabs, 3_000);
            if (newTab == 0)
                return false;

            knownTabs = Helper.GetAllExplorerTabs(mainHwnd).ToArray();
            await Task.Delay(75);
        }

        return true;
    }

    private static async Task<IReadOnlyList<ShellWindowSnapshot>> WaitForTargetsAsync(
        string[] targets,
        int beforeDefaultCount,
        IReadOnlySet<nint> baselineTopLevelWindows)
    {
        var timeoutMs = Math.Max(24_000, targets.Length * 2_500);
        var timeoutAt = Environment.TickCount64 + timeoutMs;
        var last = GetShellWindows();
        while (Environment.TickCount64 < timeoutAt)
        {
            last = GetShellWindows();
            var allTargetsFound = targets.All(target => last.Any(window => IsSameFolder(window, target)));
            var defaultCount = last.Count(IsDefaultLocation);
            var targetWindowsMerged = GetUnexpectedTargetTopLevelWindows(last, targets, baselineTopLevelWindows).Length == 0;
            if (allTargetsFound && defaultCount <= beforeDefaultCount && targetWindowsMerged)
                return last;

            await Task.Delay(200);
        }

        return last;
    }

    private static async Task<IReadOnlyList<ShellWindowSnapshot>> WaitForTargetCountAsync(string target, int expectedCount, int timeoutMs)
    {
        var timeoutAt = Environment.TickCount64 + timeoutMs;
        var last = GetShellWindows();
        while (Environment.TickCount64 < timeoutAt)
        {
            last = GetShellWindows();
            if (last.Count(window => IsSameFolder(window, target)) == expectedCount)
                return last;

            await Task.Delay(100);
        }

        return last;
    }

    private static async Task<IReadOnlyList<ShellWindowSnapshot>> WaitForStableShellWindowsAsync(int timeoutMs)
    {
        var timeoutAt = Environment.TickCount64 + timeoutMs;
        var last = GetShellWindows();
        while (Environment.TickCount64 < timeoutAt)
        {
            await Task.Delay(150);
            var current = GetShellWindows();
            if (ShellWindowSnapshotsEqual(last, current))
                return current;

            last = current;
        }

        return last;
    }

    private static bool ShellWindowSnapshotsEqual(IReadOnlyList<ShellWindowSnapshot> left, IReadOnlyList<ShellWindowSnapshot> right)
    {
        var leftKeys = left
            .Select(window => $"{window.Hwnd}|{window.TabHwnd}|{NormalizePath(window.Path)}|{window.LocationUrl}")
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var rightKeys = right
            .Select(window => $"{window.Hwnd}|{window.TabHwnd}|{NormalizePath(window.Path)}|{window.LocationUrl}")
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        return leftKeys.SequenceEqual(rightKeys, StringComparer.OrdinalIgnoreCase);
    }

    private static ShellWindowSnapshot[] GetUnexpectedTargetTopLevelWindows(
        IReadOnlyList<ShellWindowSnapshot> windows,
        string[] targets,
        IReadOnlySet<nint> baselineTopLevelWindows)
    {
        if (baselineTopLevelWindows.Count == 0)
            return [];

        return windows
            .Where(window => targets.Any(target => IsSameFolder(window, target)))
            .Where(window => !baselineTopLevelWindows.Contains((nint)window.Hwnd))
            .ToArray();
    }

    private static IReadOnlyList<ShellWindowSnapshot> GetShellWindows()
    {
        var result = new List<ShellWindowSnapshot>();
        var shellType = Type.GetTypeFromProgID("Shell.Application")
                        ?? throw new InvalidOperationException("Shell.Application COM object is not available.");

        object? shell = null;
        object? windows = null;
        try
        {
            shell = Activator.CreateInstance(shellType);
            windows = shellType.InvokeMember("Windows", System.Reflection.BindingFlags.InvokeMethod, null, shell, []);
            var count = Convert.ToInt32(windows!.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, windows, []));

            for (var i = 0; i < count; i++)
            {
                object? window = null;
                try
                {
                    window = windows.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, windows, [i]);
                    if (window == null)
                        continue;

                    result.Add(ReadShellWindow(window));
                }
                catch
                {
                    //
                }
                finally
                {
                    ReleaseComObject(window);
                }
            }
        }
        finally
        {
            ReleaseComObject(windows);
            ReleaseComObject(shell);
        }

        return result;
    }

    private static ShellWindowSnapshot ReadShellWindow(object window)
    {
        var type = window.GetType();
        var hwnd = Convert.ToInt64(type.InvokeMember("HWND", System.Reflection.BindingFlags.GetProperty, null, window, []));
        var name = Convert.ToString(type.InvokeMember("LocationName", System.Reflection.BindingFlags.GetProperty, null, window, [])) ?? string.Empty;
        var url = Convert.ToString(type.InvokeMember("LocationURL", System.Reflection.BindingFlags.GetProperty, null, window, [])) ?? string.Empty;
        var tabHwnd = GetTabHandle(window);
        var path = string.Empty;

        try
        {
            var document = type.InvokeMember("Document", System.Reflection.BindingFlags.GetProperty, null, window, []);
            var folder = document?.GetType().InvokeMember("Folder", System.Reflection.BindingFlags.GetProperty, null, document, []);
            var self = folder?.GetType().InvokeMember("Self", System.Reflection.BindingFlags.GetProperty, null, folder, []);
            path = Convert.ToString(self?.GetType().InvokeMember("Path", System.Reflection.BindingFlags.GetProperty, null, self, [])) ?? string.Empty;

            ReleaseComObject(self);
            ReleaseComObject(folder);
            ReleaseComObject(document);
        }
        catch
        {
            //
        }

        return new ShellWindowSnapshot(hwnd, name, url, path, tabHwnd);
    }
    private static long GetTabHandle(object window)
    {
        try
        {
            if (window is not ShellServiceProvider serviceProvider)
                return 0;

            var serviceGuid = ShellBrowserGuid;
            var interfaceGuid = ShellBrowserGuid;
            serviceProvider.QueryService(ref serviceGuid, ref interfaceGuid, out var shellBrowser);
            if (shellBrowser == null)
                return 0;

            try
            {
                shellBrowser.GetWindow(out var hWnd);
                return hWnd;
            }
            finally
            {
                ReleaseComObject(shellBrowser);
            }
        }
        catch
        {
            return 0;
        }
    }

    private static void CloseTestShellWindows(string root)
    {
        var shellType = Type.GetTypeFromProgID("Shell.Application");
        if (shellType == null)
            return;

        object? shell = null;
        object? windows = null;
        try
        {
            shell = Activator.CreateInstance(shellType);
            windows = shellType.InvokeMember("Windows", System.Reflection.BindingFlags.InvokeMethod, null, shell, []);
            var count = Convert.ToInt32(windows!.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, windows, []));

            for (var i = count - 1; i >= 0; i--)
            {
                object? window = null;
                try
                {
                    window = windows.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, windows, [i]);
                    if (window == null || !IsSameFolder(ReadShellWindow(window), root, allowDescendant: true))
                        continue;

                    window.GetType().InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, window, []);
                }
                catch
                {
                    //
                }
                finally
                {
                    ReleaseComObject(window);
                }
            }
        }
        finally
        {
            ReleaseComObject(windows);
            ReleaseComObject(shell);
        }
    }

    private static void CloseShellWindowByHwnd(nint hWnd)
    {
        if (hWnd == 0)
            return;

        var shellType = Type.GetTypeFromProgID("Shell.Application");
        if (shellType == null)
            return;

        object? shell = null;
        object? windows = null;
        try
        {
            shell = Activator.CreateInstance(shellType);
            windows = shellType.InvokeMember("Windows", System.Reflection.BindingFlags.InvokeMethod, null, shell, []);
            var count = Convert.ToInt32(windows!.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, windows, []));

            for (var i = count - 1; i >= 0; i--)
            {
                object? window = null;
                try
                {
                    window = windows.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, windows, [i]);
                    if (window == null)
                        continue;

                    var hwnd = Convert.ToInt64(window.GetType().InvokeMember("HWND", System.Reflection.BindingFlags.GetProperty, null, window, []));
                    if ((nint)hwnd == hWnd)
                        window.GetType().InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, window, []);
                }
                catch
                {
                    //
                }
                finally
                {
                    ReleaseComObject(window);
                }
            }
        }
        finally
        {
            ReleaseComObject(windows);
            ReleaseComObject(shell);
        }
    }

    private static void CloseShellWindowsByFolder(string folder)
    {
        var shellType = Type.GetTypeFromProgID("Shell.Application");
        if (shellType == null)
            return;

        object? shell = null;
        object? windows = null;
        try
        {
            shell = Activator.CreateInstance(shellType);
            windows = shellType.InvokeMember("Windows", System.Reflection.BindingFlags.InvokeMethod, null, shell, []);
            var count = Convert.ToInt32(windows!.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, windows, []));

            for (var i = count - 1; i >= 0; i--)
            {
                object? window = null;
                try
                {
                    window = windows.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, windows, [i]);
                    if (window == null || !IsSameFolder(ReadShellWindow(window), folder))
                        continue;

                    window.GetType().InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, window, []);
                }
                catch
                {
                    //
                }
                finally
                {
                    ReleaseComObject(window);
                }
            }
        }
        finally
        {
            ReleaseComObject(windows);
            ReleaseComObject(shell);
        }
    }

    private static void ConcealExplorerWindowAsOrphan(nint hWnd)
    {
        if (hWnd == 0)
            return;

        Helper.UpdateWindowLayered(hWnd, remove: false);
        WinApi.SetLayeredWindowAttributes(hWnd, 0, 0, WinApi.LWA_ALPHA);
        if (WinApi.GetWindowRect(hWnd, out var rect))
        {
            const uint flags = WinApi.SWP_NOSIZE | WinApi.SWP_NOZORDER | WinApi.SWP_NOACTIVATE | WinApi.SWP_FRAMECHANGED;
            WinApi.SetWindowPos(hWnd, 0, rect.Left, rect.Top, 0, 0, flags);
            WinApi.SetLayeredWindowAttributes(hWnd, 0, 0, WinApi.LWA_ALPHA);
        }
    }

    private static bool IsExplorerWindowRestored(nint hWnd)
    {
        return hWnd != 0 &&
               Helper.IsFileExplorerWindow(hWnd) &&
               WinApi.IsWindowVisible(hWnd) &&
               !IsFullyTransparent(hWnd) &&
               WinApi.GetWindowRect(hWnd, out var rect) &&
               IsOnScreen(rect);
    }

    private static void TryRestoreExplorerWindow(nint hWnd)
    {
        if (hWnd == 0)
            return;

        try
        {
            WinApi.SetLayeredWindowAttributes(hWnd, 0, 255, WinApi.LWA_ALPHA);
            Helper.UpdateWindowLayered(hWnd, remove: true);
            WinApi.ShowWindow(hWnd, WinApi.SW_SHOWNOACTIVATE);
            if (WinApi.GetWindowRect(hWnd, out var rect))
            {
                var flags = WinApi.SWP_NOSIZE | WinApi.SWP_NOZORDER | WinApi.SWP_NOACTIVATE | WinApi.SWP_SHOWWINDOW | WinApi.SWP_FRAMECHANGED;
                if (IsOnScreen(rect))
                    WinApi.SetWindowPos(hWnd, 0, rect.Left, rect.Top, 0, 0, flags);
                else
                    WinApi.SetWindowPos(hWnd, 0, 120, 120, 0, 0, flags);
            }
        }
        catch
        {
            //
        }
    }

    private static bool IsFullyTransparent(nint window)
    {
        var exStyle = WinApi.GetWindowLong(window, WinApi.GWL_EXSTYLE);
        return (exStyle & WinApi.WS_EX_LAYERED) != 0 &&
               WinApi.GetLayeredWindowAttributes(window, out _, out var alpha, out var flags) &&
               (flags & (uint)WinApi.LWA_ALPHA) != 0 &&
               alpha == 0;
    }

    private static bool IsOnScreen(RECT rect)
    {
        var width = rect.Right - rect.Left;
        var height = rect.Bottom - rect.Top;

        return width >= 80 &&
               height >= 80 &&
               rect.Right > 0 &&
               rect.Bottom > 0 &&
               rect.Left > -1_000 &&
               rect.Top > -1_000;
    }

    private static void DumpWindowVisibility(nint hWnd)
    {
        var exStyle = WinApi.GetWindowLong(hWnd, WinApi.GWL_EXSTYLE);
        var layered = (exStyle & WinApi.WS_EX_LAYERED) != 0;
        var alphaText = "n/a";
        if (layered && WinApi.GetLayeredWindowAttributes(hWnd, out _, out var alpha, out var flags))
            alphaText = $"{alpha} flags={flags}";

        var rectText = WinApi.GetWindowRect(hWnd, out var rect)
            ? $"{rect.Left},{rect.Top},{rect.Right},{rect.Bottom}"
            : "n/a";

        Console.Error.WriteLine($"HWND={hWnd} Explorer={Helper.IsFileExplorerWindow(hWnd)} Visible={WinApi.IsWindowVisible(hWnd)} Layered={layered} Alpha={alphaText} Rect={rectText}");
    }

    private static bool IsDefaultLocation(ShellWindowSnapshot window)
    {
        return window.LocationName is "This PC" or "此电脑" ||
               window.LocationUrl.Contains("20D04FE0-3AEA-1069-A2D8-08002B30309D", StringComparison.OrdinalIgnoreCase) ||
               window.Path.Contains("20D04FE0-3AEA-1069-A2D8-08002B30309D", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSameFolder(ShellWindowSnapshot window, string folder, bool allowDescendant = false)
    {
        var windowPath = NormalizePath(window.Path);
        if (string.IsNullOrWhiteSpace(windowPath) && Uri.TryCreate(window.LocationUrl, UriKind.Absolute, out var uri) && uri.IsFile)
            windowPath = NormalizePath(uri.LocalPath);

        var target = NormalizePath(folder);
        if (string.IsNullOrWhiteSpace(windowPath) || string.IsNullOrWhiteSpace(target))
            return false;

        return allowDescendant
            ? windowPath.Equals(target, StringComparison.OrdinalIgnoreCase) ||
              windowPath.StartsWith(target + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
            : windowPath.Equals(target, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizePath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return string.Empty;

        try
        {
            return Path.GetFullPath(path.Trim().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        }
        catch
        {
            return path.Trim();
        }
    }

    private static string? GetOption(string[] args, string name)
    {
        for (var i = 0; i < args.Length - 1; i++)
            if (StringComparer.OrdinalIgnoreCase.Equals(args[i], name))
                return args[i + 1];

        return null;
    }

    private static bool HasOption(string[] args, string name)
    {
        return args.Any(arg => StringComparer.OrdinalIgnoreCase.Equals(arg, name));
    }

    private static void DumpShellWindows(IReadOnlyList<ShellWindowSnapshot> windows)
    {
        foreach (var window in windows)
            Console.Error.WriteLine($"HWND={window.Hwnd} Tab={window.TabHwnd} Name='{window.LocationName}' Url='{window.LocationUrl}' Path='{window.Path}'");
    }

    private static void DumpDebugLog(string debugLog)
    {
        if (!File.Exists(debugLog))
            return;

        Console.Error.WriteLine("WinTab debug log:");
        using var stream = new FileStream(debugLog, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var reader = new StreamReader(stream);
        var lines = new List<string>();
        while (reader.ReadLine() is { } line)
            lines.Add(line);

        foreach (var line in lines.TakeLast(800))
            Console.Error.WriteLine(line);
    }

    private static void ReleaseComObject(object? value)
    {
        if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
    }

    private static nint GetActiveTabHandle(nint windowHandle)
    {
        return WinApi.FindWindowEx(windowHandle, 0, "ShellTabWindowClass", null);
    }

    private static Task<nint> WaitForActiveTabAsync(nint windowHandle, nint expectedTabHandle, int timeoutMs)
    {
        return Helper.DoUntilConditionAsync(
            () => GetActiveTabHandle(windowHandle),
            handle => handle == expectedTabHandle,
            timeoutMs,
            15);
    }

    private static Task<nint> WaitForAnyActiveTabAsync(nint windowHandle, int timeoutMs)
    {
        return Helper.DoUntilConditionAsync(
            () => GetActiveTabHandle(windowHandle),
            handle => handle != 0,
            timeoutMs,
            15);
    }

    private static async Task<bool> SelectAnyOtherTabAsync(nint windowHandle, nint currentTabHandle, int timeoutMs)
    {
        var tabs = Helper.GetAllExplorerTabs(windowHandle).ToArray();
        if (tabs.Length < 2)
            return false;

        var perTabTimeout = Math.Max(80, timeoutMs / tabs.Length);
        for (var i = 0; i < tabs.Length; i++)
        {
            WinApi.SendMessage(windowHandle, WinApi.WM_COMMAND, 0xA221, i + 1);
            var activeTab = await Helper.DoUntilConditionAsync(
                () => GetActiveTabHandle(windowHandle),
                handle => handle != 0 && handle != currentTabHandle,
                perTabTimeout,
                15);

            if (activeTab != 0 && activeTab != currentTabHandle)
                return true;
        }

        return false;
    }

    private static void TryKill(Process process)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Kill();
                process.WaitForExit(3_000);
            }
        }
        catch
        {
            //
        }
    }

    private static void KillExistingWinTabProcesses()
    {
        foreach (var existingProcess in Process.GetProcessesByName("WinTab"))
            TryKill(existingProcess);
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(
                    path,
                    Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                    Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
            }
        }
        catch
        {
            //
        }
    }

    private sealed record ShellWindowSnapshot(long Hwnd, string LocationName, string LocationUrl, string Path, long TabHwnd);

    private sealed class ExplorerProcessMonitor
    {
        private readonly int[] _processIds;
        private readonly CancellationTokenSource _cancellation = new();
        private readonly ConcurrentDictionary<int, byte> _exitedProcessIds = new();
        private Task? _pollingTask;

        private ExplorerProcessMonitor(IEnumerable<nint> explorerWindows)
        {
            _processIds = explorerWindows
                .Select(window =>
                {
                    WinApi.GetWindowThreadProcessId(window, out var pid);
                    return (int)pid;
                })
                .Where(pid => pid > 0)
                .Distinct()
                .ToArray();
        }

        public IEnumerable<int> ExitedProcessIds => _exitedProcessIds.Keys.OrderBy(pid => pid);

        public static ExplorerProcessMonitor Start(IEnumerable<nint> explorerWindows)
        {
            var monitor = new ExplorerProcessMonitor(explorerWindows);
            monitor._pollingTask = Task.Run(() => monitor.PollAsync(monitor._cancellation.Token));
            return monitor;
        }

        public async Task StopAsync()
        {
            _cancellation.Cancel();
            if (_pollingTask != null)
            {
                try
                {
                    await _pollingTask;
                }
                catch (OperationCanceledException)
                {
                    //
                }
            }

            _cancellation.Dispose();
        }

        private async Task PollAsync(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                foreach (var processId in _processIds)
                {
                    try
                    {
                        using var process = Process.GetProcessById(processId);
                        if (process.HasExited)
                            _exitedProcessIds[processId] = 0;
                    }
                    catch
                    {
                        _exitedProcessIds[processId] = 0;
                    }
                }

                await Task.Delay(100, cancellationToken);
            }
        }
    }

    private sealed class ActiveTabSequenceMonitor
    {
        private readonly nint _windowHandle;
        private readonly CancellationTokenSource _cancellation = new();
        private readonly ConcurrentQueue<nint> _observedTabs = new();
        private Task? _pollingTask;

        private ActiveTabSequenceMonitor(nint windowHandle)
        {
            _windowHandle = windowHandle;
        }

        public IReadOnlyList<nint> ObservedTabs => _observedTabs.ToArray();

        public static ActiveTabSequenceMonitor Start(nint windowHandle)
        {
            var monitor = new ActiveTabSequenceMonitor(windowHandle);
            monitor._pollingTask = Task.Run(() => monitor.PollAsync(monitor._cancellation.Token));
            return monitor;
        }

        public async Task StopAsync()
        {
            _cancellation.Cancel();
            if (_pollingTask != null)
            {
                try
                {
                    await _pollingTask;
                }
                catch (OperationCanceledException)
                {
                    //
                }
            }

            _cancellation.Dispose();
        }

        private async Task PollAsync(CancellationToken cancellationToken)
        {
            nint last = 0;
            while (!cancellationToken.IsCancellationRequested)
            {
                var activeTab = GetActiveTabHandle(_windowHandle);
                if (activeTab != 0 && activeTab != last)
                {
                    _observedTabs.Enqueue(activeTab);
                    last = activeTab;
                }

                await Task.Delay(1, cancellationToken);
            }
        }
    }

    private sealed class VisibleExplorerWindowMonitor
    {
        private static readonly TimeSpan VisibleFailureThreshold = TimeSpan.FromMilliseconds(100);
        private readonly HashSet<nint> _baselineWindows;
        private readonly CancellationTokenSource _cancellation = new();
        private readonly ConcurrentDictionary<nint, VisibleExplorerWindowSighting> _sightings = new();
        private Task? _pollingTask;

        private VisibleExplorerWindowMonitor(IEnumerable<nint> baselineWindows)
        {
            _baselineWindows = baselineWindows.ToHashSet();
        }

        public IEnumerable<VisibleExplorerWindowSighting> Sightings => _sightings.Values
            .Where(sighting => sighting.VisibleDuration >= VisibleFailureThreshold)
            .OrderBy(sighting => sighting.FirstObservedAt);

        public static VisibleExplorerWindowMonitor Start(IEnumerable<nint> baselineWindows)
        {
            var monitor = new VisibleExplorerWindowMonitor(baselineWindows);
            monitor._pollingTask = Task.Run(() => monitor.PollAsync(monitor._cancellation.Token));
            return monitor;
        }

        public async Task StopAsync()
        {
            _cancellation.Cancel();
            if (_pollingTask != null)
            {
                try
                {
                    await _pollingTask;
                }
                catch (OperationCanceledException)
                {
                    //
                }
            }

            _cancellation.Dispose();
        }

        private async Task PollAsync(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                foreach (var window in Helper.GetAllExplorerWindows())
                {
                    if (_baselineWindows.Contains(window) ||
                        !WinApi.IsWindowVisible(window) ||
                        IsFullyTransparent(window) ||
                        !WinApi.GetWindowRect(window, out var rect) ||
                        !IsOnScreen(rect))
                    {
                        continue;
                    }

                    var now = DateTime.Now;
                    _sightings.AddOrUpdate(
                        window,
                        new VisibleExplorerWindowSighting(
                            window,
                            rect.Left,
                            rect.Top,
                            rect.Right,
                            rect.Bottom,
                            now,
                            now,
                            1),
                        (_, existing) => existing.Observe(rect, now));
                }

                await Task.Delay(5, cancellationToken);
            }
        }

    }

    private sealed record VisibleExplorerWindowSighting(
        nint Hwnd,
        int Left,
        int Top,
        int Right,
        int Bottom,
        DateTime FirstObservedAt,
        DateTime LastObservedAt,
        int Observations)
    {
        public TimeSpan VisibleDuration => LastObservedAt - FirstObservedAt;

        public VisibleExplorerWindowSighting Observe(RECT rect, DateTime observedAt)
        {
            return this with
            {
                Left = rect.Left,
                Top = rect.Top,
                Right = rect.Right,
                Bottom = rect.Bottom,
                LastObservedAt = observedAt,
                Observations = Observations + 1
            };
        }
    }

    private sealed class TransientDefaultTabMonitor
    {
        // ShellWindows can expose a newly registered default tab before it is painted.
        // Persisting for a frame is the boundary this test treats as user-visible flicker.
        private static readonly TimeSpan DefaultTabFailureThreshold = TimeSpan.FromMilliseconds(16);
        private readonly HashSet<nint> _baselineWindows;
        private readonly int _baselineDefaultCount;
        private readonly CancellationTokenSource _cancellation = new();
        private readonly ConcurrentDictionary<nint, TransientDefaultTabSighting> _activeSightings = new();
        private readonly ConcurrentBag<TransientDefaultTabSighting> _confirmedSightings = [];
        private Task? _pollingTask;

        private TransientDefaultTabMonitor(IEnumerable<nint> baselineWindows, int baselineDefaultCount)
        {
            _baselineWindows = baselineWindows.ToHashSet();
            _baselineDefaultCount = baselineDefaultCount;
        }

        public IEnumerable<TransientDefaultTabSighting> Sightings => _confirmedSightings
            .Concat(_activeSightings.Values.Where(sighting => sighting.VisibleDuration >= DefaultTabFailureThreshold))
            .OrderBy(sighting => sighting.FirstObservedAt);

        public static TransientDefaultTabMonitor Start(IEnumerable<nint> baselineWindows, int baselineDefaultCount)
        {
            var monitor = new TransientDefaultTabMonitor(baselineWindows, baselineDefaultCount);
            monitor._pollingTask = Task.Run(() => monitor.PollAsync(monitor._cancellation.Token));
            return monitor;
        }

        public async Task StopAsync()
        {
            _cancellation.Cancel();
            if (_pollingTask != null)
            {
                try
                {
                    await _pollingTask;
                }
                catch (OperationCanceledException)
                {
                    //
                }
            }

            _cancellation.Dispose();
        }

        private async Task PollAsync(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                var defaultTabsByWindow = GetShellWindows()
                    .Where(window => _baselineWindows.Contains((nint)window.Hwnd))
                    .Where(IsDefaultLocation)
                    .Where(IsActiveTab)
                    .GroupBy(window => (nint)window.Hwnd)
                    .Select(group => new { Hwnd = group.Key, Count = group.Count() })
                    .ToArray();

                var currentDefaultCount = defaultTabsByWindow.Sum(item => item.Count);
                var activeWindows = defaultTabsByWindow.Select(item => item.Hwnd).ToHashSet();
                if (currentDefaultCount > _baselineDefaultCount)
                {
                    var now = DateTime.Now;
                    foreach (var item in defaultTabsByWindow)
                    {
                        _activeSightings.AddOrUpdate(
                            item.Hwnd,
                            new TransientDefaultTabSighting(item.Hwnd, item.Count, now, now, 1),
                            (_, existing) => existing.Observe(item.Count, now));
                    }
                }

                foreach (var hwnd in _activeSightings.Keys)
                {
                    if (activeWindows.Contains(hwnd))
                        continue;

                    if (_activeSightings.TryRemove(hwnd, out var completed) &&
                        completed.VisibleDuration >= DefaultTabFailureThreshold)
                    {
                        _confirmedSightings.Add(completed);
                    }
                }

                await Task.Delay(5, cancellationToken);
            }
        }

        private static bool IsActiveTab(ShellWindowSnapshot window)
        {
            var activeTab = WinApi.FindWindowEx((nint)window.Hwnd, 0, "ShellTabWindowClass", null);
            return activeTab != 0 &&
                   window.TabHwnd != 0 &&
                   activeTab == (nint)window.TabHwnd;
        }
    }

    private sealed record TransientDefaultTabSighting(
        nint Hwnd,
        int DefaultCount,
        DateTime FirstObservedAt,
        DateTime LastObservedAt,
        int Observations)
    {
        public TimeSpan VisibleDuration => LastObservedAt - FirstObservedAt;

        public TransientDefaultTabSighting Observe(int defaultCount, DateTime observedAt)
        {
            return this with
            {
                DefaultCount = Math.Max(DefaultCount, defaultCount),
                LastObservedAt = observedAt,
                Observations = Observations + 1
            };
        }
    }

    private sealed class HiddenDefaultWindowMonitor
    {
        private readonly HashSet<nint> _baselineWindows;
        private readonly CancellationTokenSource _cancellation = new();
        private readonly ConcurrentDictionary<nint, HiddenDefaultWindowSighting> _sightings = new();
        private Task? _pollingTask;

        private HiddenDefaultWindowMonitor(IEnumerable<nint> baselineWindows)
        {
            _baselineWindows = baselineWindows.ToHashSet();
        }

        public IEnumerable<HiddenDefaultWindowSighting> Sightings => _sightings.Values
            .OrderBy(sighting => sighting.FirstSeen);

        public static HiddenDefaultWindowMonitor Start(IEnumerable<nint> baselineWindows)
        {
            var monitor = new HiddenDefaultWindowMonitor(baselineWindows);
            monitor._pollingTask = Task.Run(() => monitor.PollAsync(monitor._cancellation.Token));
            return monitor;
        }

        public async Task StopAsync()
        {
            _cancellation.Cancel();
            if (_pollingTask != null)
            {
                try
                {
                    await _pollingTask;
                }
                catch (OperationCanceledException)
                {
                    //
                }
            }

            _cancellation.Dispose();
        }

        private async Task PollAsync(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                foreach (var window in GetShellWindows())
                {
                    var hWnd = (nint)window.Hwnd;
                    if (_baselineWindows.Contains(hWnd) || !IsDefaultLocation(window))
                        continue;

                    var transparent = IsFullyTransparent(hWnd);
                    var now = DateTime.Now;
                    _sightings.AddOrUpdate(
                        hWnd,
                        new HiddenDefaultWindowSighting(hWnd, now, transparent ? now : null, transparent ? now : null, transparent ? null : now, transparent ? null : now),
                        (_, existing) => existing.Observe(transparent, now));
                }

                try
                {
                    await Task.Delay(20, cancellationToken);
                }
                catch (OperationCanceledException)
                {
                    return;
                }
            }
        }
    }

    private sealed record HiddenDefaultWindowSighting(
        nint Hwnd,
        DateTime FirstSeen,
        DateTime? FirstHiddenAt,
        DateTime? LastHiddenAt,
        DateTime? FirstVisibleAt,
        DateTime? LastVisibleAt)
    {
        public double HiddenMs => FirstHiddenAt is { } start && LastHiddenAt is { } end
            ? (end - start).TotalMilliseconds
            : 0;

        public HiddenDefaultWindowSighting Observe(bool transparent, DateTime observedAt)
        {
            if (!transparent)
            {
                return this with
                {
                    FirstVisibleAt = FirstVisibleAt ?? observedAt,
                    LastVisibleAt = observedAt
                };
            }

            return this with
            {
                FirstHiddenAt = FirstHiddenAt ?? observedAt,
                LastHiddenAt = observedAt
            };
        }
    }
}
