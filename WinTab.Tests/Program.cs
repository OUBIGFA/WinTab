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
            ("closing merge sources are not restored as This PC intermediates", ExplorerTabSelectionTests.ClosingMergeSourcesAreNotRestoredAsThisPcIntermediates),
            ("Early WinEvent hide does not wait on merge target discovery", ExplorerTabSelectionTests.EarlyHideDoesNotWaitOnMergeTargetDiscovery),
            ("first run settings window fits without scrolling", ExplorerTabSelectionTests.FirstRunSettingsWindowFitsWithoutScrolling),
            ("settings resize writes are debounced and flushed", ExplorerTabSelectionTests.SettingsResizeWritesAreDebouncedAndFlushed),
            ("Explorer hot paths avoid repeated polling work", ExplorerTabSelectionTests.ExplorerHotPathsAvoidRepeatedPollingWork),
            ("window merge uses a single registration lifecycle", ExplorerTabSelectionTests.WindowMergeUsesSingleRegistrationLifecycle),
            ("update checks prefer the current architecture installer", ExplorerTabSelectionTests.UpdateCheckPrefersCurrentArchitectureInstaller),
            ("manual update checks show immediate UI feedback", ExplorerTabSelectionTests.ManualUpdateChecksShowImmediateUiFeedback),
            ("double-click close can continue after a tab-strip hit-test refresh gap", ExplorerTabDoubleClickCloseTests.ContinuousDoubleClicksCloseNextTabWithoutIntermediateClick),
            ("double-click close chain ignores points outside the double-click geometry", ExplorerTabDoubleClickCloseTests.CloseChainFallbackIgnoresDifferentPoints),
            ("TabSelectionEngine short-circuits when the target is already active", ExplorerTabSelectionTests.TabSelectionEngineAlreadyActiveReturnsImmediately),
            ("TabSelectionEngine cycles until the target becomes active", ExplorerTabSelectionTests.TabSelectionEngineCyclesUntilTargetActive),
            ("TabSelectionEngine returns false when the target never appears", ExplorerTabSelectionTests.TabSelectionEngineReturnsFalseWhenTargetMissing),
            ("TabSelectionEngine refuses a zero target without side effects", ExplorerTabSelectionTests.TabSelectionEngineRefusesZeroTarget),
            ("TabSelectionEngine reuses the existing tab without opening a duplicate", ExplorerTabSelectionTests.TabSelectionEngineDoesNotOpenNewTabsWhileCycling),
            ("TabSelectionEngine waits for the delayed active tab update without skipping", ExplorerTabSelectionTests.TabSelectionEngineWaitsForDelayedActiveUpdate),
            ("TabSelectionEngine survives transient zero active readings", ExplorerTabSelectionTests.TabSelectionEngineSurvivesTransientZeroActive),
            ("TabSelectionEngine honors the total timeout even with many tabs", ExplorerTabSelectionTests.TabSelectionEngineHonorsTotalTimeout),
            ("TabSelectionEngine does not revisit indexes during cycling", ExplorerTabSelectionTests.TabSelectionEngineDoesNotRevisitIndexes),
            ("OnWindowShown filters WinEvent sub-element notifications", ExplorerTabSelectionTests.OnWindowShownFiltersSubElementWinEvents),
            ("OnWindowShown skips non-Explorer top-level windows", ExplorerTabSelectionTests.OnWindowShownSkipsNonExplorerWindows),
            ("MergeSourceConcealPulse has an absolute ceiling", ExplorerTabSelectionTests.MergeSourceConcealPulseHasAbsoluteCeiling),
            ("MergeSourceConcealPulse never restarts itself in finally", ExplorerTabSelectionTests.MergeSourceConcealPulseNeverRestartsItselfInFinally),
            ("MergeSourceConcealPulse sleep period is at least 25 ms", ExplorerTabSelectionTests.MergeSourceConcealPulseSleepIsAtLeast25Ms),
            ("ExplorerWatcher Dispose releases the explorer-check timer", ExplorerTabSelectionTests.ExplorerWatcherDisposeReleasesExplorerCheckTimer),
            ("Pre-existing Explorer windows are not concealed during startup race", ExplorerTabSelectionTests.PreExistingExplorerWindowsAreNotConcealedDuringStartupRace),
            ("startup conceal pulse waits until pre-existing Explorer windows are protected", ExplorerTabSelectionTests.StartupConcealPulseWaitsUntilPreExistingExplorerWindowsAreProtected),
            ("Recycle Bin and third-party folder opens reuse existing tab without duplicates", ExplorerTabSelectionTests.RecycleBinAndExternalFolderOpensReuseExistingTabWithoutDuplicates),
            ("Recycle Bin and virtual folder PIDL resolution and equivalence", ExplorerTabSelectionTests.RecycleBinAndVirtualFolderPidlResolutionAndEquivalence),
            ("Tab selection budget tolerates slow shell folders without skipping cycles", ExplorerTabSelectionTests.TabSelectionBudgetToleratesSlowShellFoldersWithoutSkippingCycles),
            ("Tab search matches pre-hook windows so external opens find their tab during races", ExplorerTabSelectionTests.TabSearchMatchesPreHookWindowsSoExternalOpensFindTheirTabDuringRaces),
            ("Tab selection engine still converges when each Explorer step is slow", ExplorerTabSelectionTests.TabSelectionEngineSlowExplorerStillConvergesWithGenerousPerStepBudget),
            ("NavigateComplete2 handler signals tcs unconditionally so merges return on the first event", ExplorerTabSelectionTests.NavigateCompleteHandlerSignalsUnconditionally),
            ("NavigationCompleteWaitMs uses a short budget so merges do not stall on dropped events", ExplorerTabSelectionTests.NavigateCompleteWaitBudgetIsShort),
            ("Navigation equivalence check runs after WhenAny so the fast-path returns within a few ms of the event", ExplorerTabSelectionTests.NavigationFastPathChecksLocationAfterEvent),
            ("WaitForExplorerTabCount uses a short budget so merge decisions do not stall on a stable signal", ExplorerTabSelectionTests.WaitForExplorerTabCountUsesShortBudget),
            ("CloseMergedSourceWindowAsync uses short verification budgets so rapid detach/merge cycles do not pile up", ExplorerTabSelectionTests.CloseMergedSourceWindowUsesShortBudget),
            ("ShellWindows registration loop is bounded so rapid Explorer launches do not stack 8x150ms of idle wait", ExplorerTabSelectionTests.ShellWindowRegistrationLoopIsBounded),
            ("Race-based navigation completion returns within a couple of hundred ms when the handler fires", ExplorerTabSelectionTests.OpenTabFastPathReturnsBeforeNavigationVerificationOnComplete),
            ("Reused hwnd during _processedHWnds grace is hooked instead of being left as a transparent residual", ExplorerTabSelectionTests.ReusedHwndDuringProcessedGraceIsHookedNotLeftTransparent)
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
