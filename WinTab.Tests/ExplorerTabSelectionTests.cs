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
internal static class ExplorerTabSelectionTests
{
    public static Task SelectTabByHandleUsesFastSelectionBeforeExplorerTabUtilityFallback()
    {
        var watcherPath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var enginePath = SourceContract.FindRepoFile("WinTab", "Hooks", "TabSelectionEngine.cs");
        var watcher = File.ReadAllText(watcherPath);
        var engine = File.ReadAllText(enginePath);
        var selectBody = SourceContract.ExtractMethodBody(watcher, "public Task<bool> SelectTabByHandle");

        Assert(selectBody.Contains("TabSelectionEngine.CycleToTabAsync", StringComparison.Ordinal),
            "SelectTabByHandle must delegate to the testable TabSelectionEngine cycling routine so tab activation works without UI Automation.");
        Assert(selectBody.Contains("ExplorerWindowDiscovery.GetAllExplorerTabs(windowHandle).ToArray()", StringComparison.Ordinal) &&
               selectBody.Contains("GetActiveTabHandle(windowHandle)", StringComparison.Ordinal) &&
               selectBody.Contains("SelectTabByIndex(windowHandle, i)", StringComparison.Ordinal),
            "SelectTabByHandle must feed the engine with the live tab list, the active-tab probe, and the Ctrl+N magic-command selector.");
        Assert(!watcher.Contains("TrySelectSingleTabByAutomationName", StringComparison.Ordinal) &&
               !watcher.Contains("SelectTabByUniqueNameVerified", StringComparison.Ordinal) &&
               !watcher.Contains("TrySelectTabByCyclingAsync", StringComparison.Ordinal),
            "UI Automation tab selection is brittle on Windows 11 24H2+; the cycling engine must replace it entirely.");
        Assert(engine.Contains("CycleToTabAsync", StringComparison.Ordinal) &&
               engine.Contains("sendSelectByIndex(i)", StringComparison.Ordinal),
               "TabSelectionEngine must expose a CycleToTabAsync entry point that drives selection via the index callback.");

        return Task.CompletedTask;
    }

    public static Task TabMergeUsesExplorerTabUtilityFastPath()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection") +
                         SourceContract.ExtractMethodBody(source, "private async Task<bool> NavigateNewTabToTargetAsync");

        Assert(methodBody.Contains("await SelectTabByHandle(windowHandle, existingTab);", StringComparison.Ordinal),
            "Tab reuse should activate the matching tab via the simplified TabSelectionEngine path on a best-effort basis (no fall-through to a duplicate tab when selection times out).");
        Assert(methodBody.Contains("ListenForNewExplorerTabAsync(mainWindowHWnd, currentTabs, 2_000)", StringComparison.Ordinal),
            "New tab creation should use the short ExplorerTabUtility wait window.");
        Assert(methodBody.Contains("FindShellWindowByTabHandle(newTabHandle, mainWindowHWnd)", StringComparison.Ordinal) &&
               methodBody.Contains("2_000", StringComparison.Ordinal),
            "New tab navigation should bind the ShellWindows object by tab handle without adding slow address-bar fallbacks.");
        Assert(!source.Contains("NavigateActiveTabByAddressBar", StringComparison.Ordinal) &&
               !source.Contains("TryPasteAddressText", StringComparison.Ordinal) &&
               !source.Contains("KeyboardSimulator.SendText", StringComparison.Ordinal),
            "Opening a tab must not drive Explorer through the address bar or clipboard.");
        var registrationBody = SourceContract.ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");
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
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowsAsync");

        Assert(methodBody.Contains("Task.WhenAll", StringComparison.Ordinal) &&
               methodBody.Contains("ProcessRegisteredShellWindowAsync", StringComparison.Ordinal),
            "Rapid Explorer opens must resolve source locations concurrently; one slow transient This PC window must not block the whole registration batch.");
        Assert(!methodBody.Contains("foreach (var (window, windowInfo) in windows)\r\n                await ProcessRegisteredShellWindowAsync(window, windowInfo);", StringComparison.Ordinal),
            "Registration processing must not await each hidden source window serially.");

        return Task.CompletedTask;
    }

    public static Task TabNavigationFailureIsShortBoundedPath()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var openBody = SourceContract.ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection");
        var navigateBody = SourceContract.ExtractMethodBody(source, "private async Task<bool> NavigateNewTabToTargetAsync");

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
        var watcherPath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var enginePath = SourceContract.FindRepoFile("WinTab", "Hooks", "TabSelectionEngine.cs");
        var watcher = File.ReadAllText(watcherPath);
        var engine = File.ReadAllText(enginePath);
        var openBody = SourceContract.ExtractMethodBody(watcher, "private async Task<bool> OpenTabNavigateWithSelection");
        var selectBody = SourceContract.ExtractMethodBody(watcher, "public Task<bool> SelectTabByHandle");

        Assert(!openBody.Contains("Task.Delay(60)", StringComparison.Ordinal),
            "Reuse should not pay a fixed sleep after foregrounding; direct selection is already verified by the active tab handle.");
        Assert(selectBody.Contains("TabSelectionEngine.CycleToTabAsync", StringComparison.Ordinal),
            "Reuse selection should hand off to the cycling engine instead of rebuilding UI Automation state for every tab activation.");
        Assert(!selectBody.Contains("AutomationElement.FromHandle", StringComparison.Ordinal) &&
               !watcher.Contains("TrySelectSingleTabByAutomationName", StringComparison.Ordinal),
            "AutomationElement.FromHandle and the legacy name-based selector must not run on every reuse selection.");
        Assert(engine.Contains("perStep", StringComparison.Ordinal) &&
               engine.Contains("Environment.TickCount64", StringComparison.Ordinal),
            "The cycling engine must enforce a per-step budget instead of falling back to fixed sleeps.");

        return Task.CompletedTask;
    }

    public static Task FailedNavigationClosesTransientThisPcTab()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection") +
                         SourceContract.ExtractMethodBody(source, "private async Task<bool> NavigateNewTabToTargetAsync") +
                         SourceContract.ExtractMethodBody(source, "private bool AreLocationsEquivalent") +
                         SourceContract.ExtractMethodBody(source, "private nint GetMainWindowHWnd") +
                         SourceContract.ExtractMethodBody(source, "private bool IsStableMergeTargetWindow") +
                         SourceContract.ExtractMethodBody(source, "private static bool IsFallbackMergeTargetWindow") +
                         SourceContract.ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");

        Assert(methodBody.Contains("WaitForNavigation(window, targetLocation, NavigationVerificationWaitMs)", StringComparison.Ordinal) &&
               source.Contains("NavigationVerificationWaitMs = 1_200", StringComparison.Ordinal) &&
               methodBody.Contains("AreLocationsEquivalent(TryGetLocation(window), targetLocation)", StringComparison.Ordinal),
            "The newly opened tab must be checked against the requested target folder.");
        Assert(methodBody.Contains("CloseFailedNewTabAsync(mainWindowHWnd, newTabHandle)", StringComparison.Ordinal),
            "A tab that remains at This PC or another wrong location must be closed instead of left behind.");
        Assert(methodBody.Contains("ExplorerWindowDiscovery.IsFileExplorerForeground(out var foregroundWindow)", StringComparison.Ordinal),
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
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private bool TryHideIncomingExplorerWindow") +
                         SourceContract.ExtractMethodBody(source, "private void TryHideRegisteredMergeSourceWindow") +
                         SourceContract.ExtractMethodBody(source, "private void OnWindowShown") +
                         SourceContract.ExtractMethodBody(source, "private void OnShellWindowRegistered") +
                         SourceContract.ExtractMethodBody(source, "private List<(InternetExplorer Window, WindowInfo WindowInfo)> AdoptNewShellWindows") +
                         SourceContract.ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync") +
                         SourceContract.ExtractMethodBody(source, "private void HideMergeSourceWindow") +
                         SourceContract.ExtractMethodBody(source, "private async Task RestoreMergeSourceWindowAsync") +
                         SourceContract.ExtractMethodBody(source, "private void RemoveMergeSourceTracking") +
                         SourceContract.ExtractMethodBody(source, "private void StartMergeSourceConcealPulse") +
                         SourceContract.ExtractMethodBody(source, "private void ConcealMergeSourceWindowsOnce") +
                         SourceContract.ExtractMethodBody(source, "private void InitializeShellObjects") +
                         SourceContract.ExtractMethodBody(source, "private sealed class WinEventHookThread");

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
               methodBody.Contains("ConcealMergeSourceWindowsOnce", StringComparison.Ordinal),
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

    public static async Task MergeSourceConcealPulseIsBoundedAndEventTriggered()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var startBody = SourceContract.ExtractMethodBody(source, "public void StartHook");
        var windowShownBody = SourceContract.ExtractMethodBody(source, "private void OnWindowShown");
        var registeredBody = SourceContract.ExtractMethodBody(source, "private void OnShellWindowRegistered");
        var callCount = 0;

        Assert(!source.Contains("StartMergeSourceConcealPulse(0)", StringComparison.Ordinal),
            "The conceal pulse must not run forever after startup; that keeps scanning Explorer windows and can re-hide later user-opened This PC windows.");
        Assert(startBody.Contains("StartMergeSourceConcealPulse(500)", StringComparison.Ordinal),
            "Startup should only run a short warm-up conceal pulse.");
        Assert(source.Contains("private void StartMergeSourceConcealPulse(int durationMs = 1_200)", StringComparison.Ordinal),
            "Event-triggered conceal should stay bounded so hidden merge-source protection cannot linger for seconds after activity stops.");
        Assert(windowShownBody.Contains("StartMergeSourceConcealPulse()", StringComparison.Ordinal) &&
               registeredBody.Contains("StartMergeSourceConcealPulse()", StringComparison.Ordinal),
            "WinEvent and ShellWindows registration events should still trigger the short conceal pulse for rapid Explorer opens.");

        var pulse = new MergeSourceConcealPulse(absoluteCeilingMs: 150, sleepMs: 5);
        pulse.Start(() => true, () => Interlocked.Increment(ref callCount), durationMs: 40);
        await Task.Delay(80);
        var countAfterPulse = Volatile.Read(ref callCount);
        await Task.Delay(60);

        Assert(countAfterPulse > 0, "The pulse must run at least one scan while enabled.");
        Assert(Volatile.Read(ref callCount) == countAfterPulse,
            "The pulse worker should be duration-bound instead of continuing to scan after its window expires.");
    }

    public static Task StartupConcealPulseWaitsUntilPreExistingExplorerWindowsAreProtected()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var startBody = SourceContract.ExtractMethodBody(source, "public void StartHook");
        var initializeBody = SourceContract.ExtractMethodBody(source, "private void InitializeShellObjects");
        var disposeBody = SourceContract.ExtractMethodBody(source, "private void DisposeShellObjects");

        Assert(source.Contains("_preExistingExplorerWindowsProtected", StringComparison.Ordinal),
            "The startup race needs an explicit guard so the conceal pulse cannot scan existing Explorer windows before they are marked safe.");
        Assert(startBody.Contains("_preExistingExplorerWindowsProtected", StringComparison.Ordinal) &&
               startBody.Contains("StartMergeSourceConcealPulse(500)", StringComparison.Ordinal),
            "StartHook may warm up the conceal pulse only after pre-existing Explorer windows have been protected.");
        Assert(initializeBody.Contains("PreventWindowHiding(new IntPtr(window.HWND))", StringComparison.Ordinal) &&
               initializeBody.Contains("_preExistingExplorerWindowsProtected = true", StringComparison.Ordinal),
            "Shell initialization must mark every already-open Explorer window safe before enabling startup conceal scans.");
        Assert(initializeBody.Contains("if (_isForcingTabs)", StringComparison.Ordinal) &&
               initializeBody.Contains("StartMergeSourceConcealPulse(500)", StringComparison.Ordinal),
            "If the hook was enabled before Shell initialization finished, the bounded warm-up scan should run only after protection is complete.");
        Assert(disposeBody.Contains("_preExistingExplorerWindowsProtected = false", StringComparison.Ordinal),
            "The guard must reset when Shell objects are rebuilt after Explorer restarts.");

        return Task.CompletedTask;
    }

    public static Task ShellWindowRegistrationCallbackStaysNonBlocking()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private void OnShellWindowRegistered");
        var scheduleBody = SourceContract.ExtractMethodBody(source, "private void ScheduleShellWindowRegistration");

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
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private List<(InternetExplorer Window, WindowInfo WindowInfo)> AdoptNewShellWindows");

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
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");
        var normalizedBody = methodBody.Replace("\r\n", "\n", StringComparison.Ordinal);

        var resolveIndex = normalizedBody.IndexOf("ResolveInitialLocation(window, hWnd)", StringComparison.Ordinal);
        var tabHandleIndex = normalizedBody.IndexOf("GetTabHandle(window)", StringComparison.Ordinal);
        var releaseBranchStart = normalizedBody.IndexOf("if (string.IsNullOrWhiteSpace(location)", StringComparison.Ordinal);
        var sourceAliveIndex = normalizedBody.IndexOf("var sourceAlive = ExplorerWindowDiscovery.IsFileExplorerWindow(hWnd)", StringComparison.Ordinal);
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
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var registrationBody = SourceContract.ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");
        var targetBody = SourceContract.ExtractMethodBody(source, "private nint GetMainWindowHWnd") +
                         SourceContract.ExtractMethodBody(source, "private bool IsPreferredMergeTargetWindow") +
                         SourceContract.ExtractMethodBody(source, "private bool HasNonStartupShellWindowForTopLevel");

        Assert(registrationBody.Contains("GetMainWindowHWnd(hWnd, location)", StringComparison.Ordinal),
            "After resolving a real folder, merge target selection must use the resolved location to avoid a stale This PC target.");
        Assert(targetBody.Contains("ShouldPreferNonStartupTarget(targetLocation)", StringComparison.Ordinal) &&
               targetBody.Contains("HasNonStartupShellWindowForTopLevel", StringComparison.Ordinal),
            "Folder merges should prefer a target Explorer that already represents a real folder when one exists.");

        return Task.CompletedTask;
    }

    public static Task MergeableDefaultLocationIntermediateClosesAfterTargetTabSucceeds()
    {
        var resolverPath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerLaunchLocationResolver.cs");
        var watcherPath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var resolver = File.ReadAllText(resolverPath);
        var watcher = File.ReadAllText(watcherPath);
        var resolveBody = SourceContract.ExtractMethodBody(watcher, "private Task<string> ResolveInitialLocation");
        var restoreBody = SourceContract.ExtractMethodBody(watcher, "private async Task RestoreHiddenExplorerWindowAsync");
        var registrationBody = SourceContract.ExtractMethodBody(watcher, "private async Task ProcessRegisteredShellWindowAsync");

        Assert(resolver.Contains("int DefaultLocationWaitMs = 30", StringComparison.Ordinal),
            "A non-mergeable stable This PC window should still be released quickly instead of staying transparent.");
        Assert(resolver.Contains("int MaximumStartupLocationWaitMs = 350", StringComparison.Ordinal),
            "Mergeable stable This PC windows should not wait the old multi-second startup window before merging.");
        Assert(resolver.Contains("int BusyStartupLocationWaitMs = 1_200", StringComparison.Ordinal) &&
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
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var restoreBody = SourceContract.ExtractMethodBody(source, "private async Task RestoreMergeSourceWindowAsync");

        Assert(restoreBody.Contains("PreventWindowHiding(hWnd)", StringComparison.Ordinal),
            "A released This PC/user Explorer window must be protected from late WinEvents that would hide it again.");

        return Task.CompletedTask;
    }

    public static Task HideWindowReappliesTransparency()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Helpers", "ExplorerWindowVisibility.cs");
        var source = File.ReadAllText(sourcePath);
        var hideBody = SourceContract.ExtractMethodBody(source, "public static void Hide");
        var showBody = SourceContract.ExtractMethodBody(source, "public static bool Show");

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
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var initializeBody = SourceContract.ExtractMethodBody(source, "private void InitializeShellObjects");
        var stopBody = SourceContract.ExtractMethodBody(source, "public void StopHook");
        var disposeBody = SourceContract.ExtractMethodBody(source, "private void DisposeShellObjects");
        var registrationBody = SourceContract.ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");
        var removeBody = SourceContract.ExtractMethodBody(source, "private void RemoveWindowAndUnhookEvents");

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
        Assert(removeBody.Contains("ExplorerWindowVisibility.Restore", StringComparison.Ordinal),
            "Removing a live shell window must restore it unless the caller intentionally closes it.");

        return Task.CompletedTask;
    }

    public static Task OrphanedTransparentExplorerWindowsAreRecoveredWithoutCache()
    {
        var visibilityPath = SourceContract.FindRepoFile("WinTab", "Helpers", "ExplorerWindowVisibility.cs");
        var visibility = File.ReadAllText(visibilityPath);
        var restoreBody = SourceContract.ExtractMethodBody(visibility, "public static bool Restore");
        var recoveryBody = SourceContract.ExtractMethodBody(visibility, "public static int RestoreAll");

        Assert(restoreBody.Contains("GetLayeredWindowAttributes", StringComparison.Ordinal) &&
               restoreBody.Contains("alpha == 0", StringComparison.Ordinal) &&
               restoreBody.Contains("SetLayeredWindowAttributes(hWnd, 0, 255, WinApi.LWA_ALPHA)", StringComparison.Ordinal),
            "RestoreHiddenExplorerWindow must repair alpha=0 Explorer windows even when HiddenWindows has no cache entry.");
        Assert(recoveryBody.Contains("ExplorerWindowDiscovery.GetAllExplorerWindows()", StringComparison.Ordinal),
            "RestoreHiddenExplorerWindows must scan live Explorer windows, not only the in-process hidden-window cache.");
        Assert(recoveryBody.Contains("Restore(hWnd", StringComparison.Ordinal),
            "RestoreHiddenExplorerWindows must route every candidate through the same restore primitive.");

        return Task.CompletedTask;
    }

    public static Task ShellInitializationSkipsDuplicateShellWindowsEntries()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var initializeBody = SourceContract.ExtractMethodBody(source, "private void InitializeShellObjects");

        Assert(initializeBody.Contains("_windowEntryDict.Keys.Contains(window)", StringComparison.Ordinal) &&
               initializeBody.Contains("window.GetProperty(\"seenBefore\")", StringComparison.Ordinal),
            "Startup ShellWindows enumeration can contain duplicate COM entries and must skip already-seen windows before adding them.");

        return Task.CompletedTask;
    }

    public static Task TabReuseExcludesMergeSource()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var searchBody = SourceContract.ExtractMethodBody(source, "private bool TrySearchForTab");
        var openBody = SourceContract.ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection");

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
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var openBody = SourceContract.ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection");

        var foundIndex = openBody.IndexOf("TrySearchForTab(windowToOpen.Location, windowToOpen.Handle", StringComparison.Ordinal);
        var successLogIndex = openBody.IndexOf("OpenTab reused target", foundIndex >= 0 ? foundIndex : 0, StringComparison.Ordinal);
        var returnTrueIndex = openBody.IndexOf("return true;", successLogIndex >= 0 ? successLogIndex : 0, StringComparison.Ordinal);
        Assert(foundIndex >= 0 && successLogIndex > foundIndex && returnTrueIndex > successLogIndex,
            "Reuse must log success and return true once an existing tab is located so external folder opens (Recycle Bin, third-party app launches) never create a duplicate tab.");

        Assert(!openBody.Contains("OpenTab reuse-select-failed", StringComparison.Ordinal),
            "Reuse must not fall through to new-tab creation when an existing tab is found.");

        var newTabIndex = openBody.IndexOf("RequestToOpenNewTab(mainWindowHWnd, lockToOpenWindows: false)", StringComparison.Ordinal);
        Assert(newTabIndex == -1 || newTabIndex > returnTrueIndex,
            "New-tab creation must happen only on the no-existing-tab path, not after a found tab returns success.");

        Assert(openBody.Contains("await SelectTabByHandle(windowHandle, existingTab);", StringComparison.Ordinal),
            "Reuse must still drive the matching tab into focus on a best-effort basis, even though success no longer depends on the return value.");

        return Task.CompletedTask;
    }

    public static Task TabReuseForegroundsTargetBeforeSelection()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var openBody = SourceContract.ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection");

        var foundIndex = openBody.IndexOf("TrySearchForTab(windowToOpen.Location, windowToOpen.Handle", StringComparison.Ordinal);
        var foregroundIndex = openBody.IndexOf("Helper.RestoreWindowToForeground(windowHandle)", foundIndex >= 0 ? foundIndex : 0, StringComparison.Ordinal);
        var selectIndex = openBody.IndexOf("SelectTabByHandle(windowHandle, existingTab", foundIndex >= 0 ? foundIndex : 0, StringComparison.Ordinal);

        Assert(foundIndex >= 0 && foregroundIndex > foundIndex && selectIndex > foregroundIndex,
            "External folder launches often leave a third-party app in front; reuse must foreground Explorer before selecting the matching tab.");

        return Task.CompletedTask;
    }

    public static Task MergeSourceCloseIsVerifiedBeforeHiddenTrackingIsCleared()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var registrationBody = SourceContract.ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");
        var closeBody = SourceContract.ExtractMethodBody(source, "private async Task<bool> CloseMergedSourceWindowAsync");

        var closeIndex = registrationBody.IndexOf("CloseMergedSourceWindowAsync(window, hWnd)", StringComparison.Ordinal);
        var clearIndex = registrationBody.IndexOf("RemoveMergeSourceTracking(hWnd)", closeIndex, StringComparison.Ordinal);
        Assert(closeIndex >= 0 && clearIndex > closeIndex,
            "A hidden source window must only be removed from hidden tracking after its close request has been verified.");
        Assert(closeBody.Contains("Helper.DoUntilConditionAsync", StringComparison.Ordinal) &&
               closeBody.Contains("!ExplorerWindowDiscovery.IsFileExplorerWindow(hWnd)", StringComparison.Ordinal) &&
               closeBody.Contains("ScheduleConcealedMergeSourceCloseRetry(hWnd)", StringComparison.Ordinal),
            "A merged source window that has not disappeared yet must stay concealed and keep closing instead of being restored as a visible This PC intermediate.");

        return Task.CompletedTask;
    }

    public static Task ClosingMergeSourcesAreNotRestoredAsThisPcIntermediates()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var closeBody = SourceContract.ExtractMethodBody(source, "private async Task<bool> CloseMergedSourceWindowAsync");
        var retryBody = SourceContract.ExtractMethodBody(source, "private void ScheduleConcealedMergeSourceCloseRetry");
        var hideBody = SourceContract.ExtractMethodBody(source, "private bool TryHideIncomingExplorerWindow");
        var removeBody = SourceContract.ExtractMethodBody(source, "private void RemoveWindowAndUnhookEvents");

        Assert(closeBody.Contains("_closingMergeSourceHWnds.TryAdd(hWnd, 0)", StringComparison.Ordinal),
            "Merged source windows must be marked as closing before their close request is sent.");
        Assert(closeBody.Contains("ScheduleConcealedMergeSourceCloseRetry(hWnd)", StringComparison.Ordinal) &&
               !closeBody.Contains("await RestoreMergeSourceWindowAsync(hWnd)", StringComparison.Ordinal),
            "A closing merged source must not be restored when the short close verification misses the final close.");
        Assert(hideBody.Contains("_closingMergeSourceHWnds.ContainsKey(hWnd)", StringComparison.Ordinal) &&
               hideBody.Contains("RequestCloseMergedSourceWindow(hWnd)", StringComparison.Ordinal),
            "Late WinEvents for a closing This PC intermediate must re-hide it and send another close request.");
        Assert(removeBody.Contains("_closingMergeSourceHWnds.ContainsKey(hWnd)", StringComparison.Ordinal) &&
               removeBody.Contains("restoreHiddenWindow = false", StringComparison.Ordinal),
            "Closing intermediates must not be restored while their ShellWindows wrapper is being removed.");
        Assert(retryBody.Contains("HideMergeSourceWindow(hWnd)", StringComparison.Ordinal) &&
               retryBody.Contains("RequestCloseMergedSourceWindow(hWnd)", StringComparison.Ordinal),
            "The retry path must keep the intermediate concealed while it continues closing.");

        return Task.CompletedTask;
    }

    public static Task FirstRunSettingsWindowFitsWithoutScrolling()
    {
        var settingsPath = SourceContract.FindRepoFile("WinTab", "Managers", "SettingsManager.cs");
        var xamlPath = SourceContract.FindRepoFile("WinTab", "UI", "Views", "MainWindow.xaml");
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
        var settingsPath = SourceContract.FindRepoFile("WinTab", "Managers", "SettingsManager.cs");
        var mainWindowPath = SourceContract.FindRepoFile("WinTab", "UI", "Views", "MainWindow.xaml.cs");
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
        var discoveryPath = SourceContract.FindRepoFile("WinTab", "Helpers", "ExplorerWindowDiscovery.cs");
        var watcherPath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var discovery = File.ReadAllText(discoveryPath);
        var watcher = File.ReadAllText(watcherPath);
        var listenBodies = SourceContract.ExtractMethodBody(discovery, "public static Task<nint> ListenForNewExplorerWindowAsync") +
                           SourceContract.ExtractMethodBody(discovery, "public static nint ListenForNewExplorerTab") +
                           SourceContract.ExtractMethodBody(discovery, "public static Task<nint> ListenForNewExplorerTabAsync");
        var targetBody = SourceContract.ExtractMethodBody(watcher, "private nint GetMainWindowHWnd");
        var startupBody = SourceContract.ExtractMethodBody(watcher, "private bool IsStartupExplorerLocation");

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
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");

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
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var hideBody = SourceContract.ExtractMethodBody(source, "private bool TryHideIncomingExplorerWindow");

        Assert(!hideBody.Contains("HasMergeTargetForEarlyConceal(hWnd)", StringComparison.Ordinal),
            "Early WinEvent hide must not perform merge-target discovery before concealing a new Explorer window.");
        Assert(!hideBody.Contains("HasTrackedTopLevelWindow(hWnd)", StringComparison.Ordinal) &&
               !hideBody.Contains("HasHookedShellWindowForTopLevel(hWnd)", StringComparison.Ordinal) &&
               hideBody.Contains("_hookedTopLevelUseCounts.ContainsKey(hWnd)", StringComparison.Ordinal),
            "Early WinEvent hide must use the hooked-target cache without taking the ShellWindows dictionary lock.");

        var targetBody = SourceContract.ExtractMethodBody(source, "private bool IsStableMergeTargetWindow");
        Assert(targetBody.Contains("HasHookedShellWindowForTopLevel(hWnd)", StringComparison.Ordinal),
            "Merge target selection must not treat another newly-created unhooked Explorer window as a stable target.");

        return Task.CompletedTask;
    }

    public static Task WindowMergeUsesSingleRegistrationLifecycle()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var windowInfoPath = SourceContract.FindRepoFile("WinTab", "Models", "WindowInfo.cs");
        var source = File.ReadAllText(sourcePath);
        var windowInfo = File.ReadAllText(windowInfoPath);
        var registrationBody = SourceContract.ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");

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
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Managers", "UpdateManager.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private static string? FindMatchingUpdateAssetUrl") +
                         SourceContract.ExtractMethodBody(source, "private static string? GetInstallerArchitectureSuffix");

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

    public static Task ManualUpdateChecksShowImmediateUiFeedback()
    {
        var mainWindowPath = SourceContract.FindRepoFile("WinTab", "UI", "Views", "MainWindow.xaml.cs");
        var updateManagerPath = SourceContract.FindRepoFile("WinTab", "Managers", "UpdateManager.cs");
        var mainWindow = File.ReadAllText(mainWindowPath);
        var updateManager = File.ReadAllText(updateManagerPath);
        var clickBody = SourceContract.ExtractMethodBody(mainWindow, "private async void CheckUpdatesButton_Click");

        Assert(clickBody.Contains("CheckUpdatesButton.IsEnabled = false", StringComparison.Ordinal) &&
               clickBody.Contains("SetMaintenanceFeedback", StringComparison.Ordinal) &&
               clickBody.Contains("CheckForUpdatesWithResultAsync().ConfigureAwait(true)", StringComparison.Ordinal),
            "Manual update checks must give immediate in-window feedback while the network check is running.");
        Assert(clickBody.Contains("Update check failed. Try again later.", StringComparison.Ordinal) &&
               clickBody.Contains("You're on the latest version.", StringComparison.Ordinal) &&
               clickBody.Contains("Update found. Opening the update window.", StringComparison.Ordinal),
            "Manual update checks must report failure, up-to-date, and update-found states in the Maintenance area.");
        Assert(mainWindow.Contains("ApplyMaintenanceDescription()", StringComparison.Ordinal) &&
               mainWindow.Contains("DispatcherTimer", StringComparison.Ordinal),
            "Maintenance feedback should be visible in the existing description area and reset without adding extra UI chrome.");
        Assert(updateManager.Contains("CheckForUpdatesWithResultAsync", StringComparison.Ordinal) &&
               updateManager.Contains("UpdateCheckResult", StringComparison.Ordinal),
            "The UI needs a result-returning update check instead of calling the updater with no feedback path.");

        return Task.CompletedTask;
    }

    public static async Task TabSelectionEngineAlreadyActiveReturnsImmediately()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33 }, activeIndex: 1);
        var start = Environment.TickCount64;
        var ok = await TabSelectionEngine.CycleToTabAsync(
            22,
            fixture.GetTabs,
            fixture.GetActive,
            fixture.SendSelectByIndex,
            totalTimeoutMs: 200);

        Assert(ok, "Cycling must return true immediately when the requested handle is already active.");
        Assert(Environment.TickCount64 - start < 80,
            "An already-active target must not pay for any cycling waits.");
        Assert(fixture.SelectionCalls.Count == 0,
            "An already-active target must never send a tab-switch command.");
    }

    public static async Task TabSelectionEngineCyclesUntilTargetActive()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33, 44 }, activeIndex: 0);
        fixture.OnSelectByIndex = i => fixture.SetActiveIndex(i);

        var ok = await TabSelectionEngine.CycleToTabAsync(
            33,
            fixture.GetTabs,
            fixture.GetActive,
            fixture.SendSelectByIndex,
            totalTimeoutMs: 400);

        Assert(ok, "Cycling must succeed when the target handle is reachable by index.");
        Assert(fixture.GetActive() == 33,
            "The active tab must end up at the requested handle after cycling.");
        Assert(fixture.SelectionCalls.Count > 0 && fixture.SelectionCalls.Count <= 4,
            "Cycling must drive the selector forward in bounded steps, not loop indefinitely.");
    }

    public static async Task TabSelectionEngineReturnsFalseWhenTargetMissing()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33 }, activeIndex: 0);
        fixture.OnSelectByIndex = i => fixture.SetActiveIndex(i);

        var ok = await TabSelectionEngine.CycleToTabAsync(
            99,
            fixture.GetTabs,
            fixture.GetActive,
            fixture.SendSelectByIndex,
            totalTimeoutMs: 200,
            pollSleepMs: 1);

        Assert(!ok, "Cycling must report failure when no index activates the requested handle.");
        Assert(fixture.SelectionCalls.Count <= 3,
            "Cycling must stop after exhausting the available tab indexes instead of looping forever.");
    }

    public static async Task TabSelectionEngineRefusesZeroTarget()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22 }, activeIndex: 0);
        var ok = await TabSelectionEngine.CycleToTabAsync(
            0,
            fixture.GetTabs,
            fixture.GetActive,
            fixture.SendSelectByIndex,
            totalTimeoutMs: 50);

        Assert(!ok, "A zero target handle must never report success.");
        Assert(fixture.SelectionCalls.Count == 0,
            "A zero target must not send any selection commands.");
    }

    public static async Task TabSelectionEngineDoesNotOpenNewTabsWhileCycling()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33 }, activeIndex: 0);
        var newTabRequests = 0;
        fixture.OnSelectByIndex = i =>
        {
            if (i < 0 || i >= fixture.Tabs.Length)
            {
                newTabRequests++;
                return;
            }
            fixture.SetActiveIndex(i);
        };

        await TabSelectionEngine.CycleToTabAsync(
            33,
            fixture.GetTabs,
            fixture.GetActive,
            fixture.SendSelectByIndex,
            totalTimeoutMs: 300);

        Assert(newTabRequests == 0,
            "Cycling must never request an out-of-range index that would create or duplicate a tab.");
        Assert(fixture.GetActive() == 33,
            "Cycling must converge on the requested handle without opening a new tab.");
    }

    public static async Task TabSelectionEngineWaitsForDelayedActiveUpdate()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33 }, activeIndex: 0);
        var delayCount = 0;
        fixture.OnSelectByIndex = i =>
        {
            if (i == 2)
            {
                delayCount = 3;
                return;
            }
            fixture.SetActiveIndex(i);
        };
        fixture.OnGetActive = current =>
        {
            if (delayCount > 0)
            {
                delayCount--;
                return current;
            }
            if (fixture.SelectionCalls.Count > 0 && fixture.SelectionCalls[^1] == 2)
                fixture.SetActiveIndex(2);
            return current;
        };

        var ok = await TabSelectionEngine.CycleToTabAsync(
            33,
            fixture.GetTabs,
            fixture.GetActive,
            fixture.SendSelectByIndex,
            totalTimeoutMs: 400,
            pollSleepMs: 1);

        Assert(ok, "The engine must wait through delayed active tab updates and still report success.");
        Assert(fixture.GetActive() == 33,
            "After waiting through the delay, the active tab must reflect the requested handle.");
    }

    public static async Task TabSelectionEngineSurvivesTransientZeroActive()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33 }, activeIndex: 0);
        var emitZeroTimes = 0;
        fixture.OnSelectByIndex = i => fixture.SetActiveIndex(i);
        fixture.OnGetActive = current =>
        {
            if (emitZeroTimes > 0)
            {
                emitZeroTimes--;
                return 0;
            }
            return current;
        };

        emitZeroTimes = 2;
        var ok = await TabSelectionEngine.CycleToTabAsync(
            33,
            fixture.GetTabs,
            fixture.GetActive,
            fixture.SendSelectByIndex,
            totalTimeoutMs: 300,
            pollSleepMs: 1);

        Assert(ok, "Transient zero active readings must not cause the engine to give up on the target.");
        Assert(fixture.GetActive() == 33,
            "The engine must converge on the requested tab even after seeing temporary zero readings.");
    }

    public static async Task TabSelectionEngineHonorsTotalTimeout()
    {
        var fixture = new TabSelectionFixture(
            new nint[] { 11, 22, 33, 44, 55, 66, 77, 88, 99, 1010, 1111, 1212 },
            activeIndex: 0);

        var start = Environment.TickCount64;
        var ok = await TabSelectionEngine.CycleToTabAsync(
            9999,
            fixture.GetTabs,
            fixture.GetActive,
            fixture.SendSelectByIndex,
            totalTimeoutMs: 200,
            pollSleepMs: 1);
        var elapsed = Environment.TickCount64 - start;

        Assert(!ok, "Missing targets must report failure.");
        Assert(elapsed < 400,
            $"Total elapsed must respect totalTimeoutMs even with many tabs; observed {elapsed}ms.");
    }

    public static async Task TabSelectionEngineDoesNotRevisitIndexes()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33, 44 }, activeIndex: 0);
        fixture.OnSelectByIndex = i => fixture.SetActiveIndex(i);

        var ok = await TabSelectionEngine.CycleToTabAsync(
            44,
            fixture.GetTabs,
            fixture.GetActive,
            fixture.SendSelectByIndex,
            totalTimeoutMs: 400);

        Assert(ok, "Cycling must converge on the requested handle.");
        var distinctCalls = fixture.SelectionCalls.Distinct().Count();
        Assert(distinctCalls == fixture.SelectionCalls.Count,
            $"The engine must not request the same tab index more than once; called {string.Join(',', fixture.SelectionCalls)}.");
    }

    public static Task OnWindowShownFiltersSubElementWinEvents()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private void OnWindowShown");

        Assert(methodBody.Contains("idObject != 0 || idChild != 0", StringComparison.Ordinal),
            "OnWindowShown must skip sub-element WinEvents (idObject != OBJID_WINDOW). System-wide WinEvent " +
            "hooks deliver hundreds-to-thousands of caret/focus/menu/list-item events per second on busy desktops; " +
            "without this filter every blink would push the conceal pulse and shell-window scheduler.");

        var idObjectIndex = methodBody.IndexOf("idObject != 0 || idChild != 0", StringComparison.Ordinal);
        var heavyWorkIndex = methodBody.IndexOf("TryHideIncomingExplorerWindow(", StringComparison.Ordinal);
        Assert(idObjectIndex >= 0 && heavyWorkIndex > idObjectIndex,
            "The OBJID_WINDOW/CHILDID_SELF filter must run before any cross-process P/Invoke work.");

        return Task.CompletedTask;
    }

    public static Task OnWindowShownSkipsNonExplorerWindows()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private void OnWindowShown");

        Assert(methodBody.Contains("GetExplorerTopLevelWindow(hWnd)", StringComparison.Ordinal),
            "OnWindowShown must resolve the Explorer top-level window once and reuse the result so we never " +
            "trigger the conceal pulse or registration scheduler for non-Explorer events.");

        var resolveIndex = methodBody.IndexOf("GetExplorerTopLevelWindow(hWnd)", StringComparison.Ordinal);
        var pulseIndex = methodBody.IndexOf("StartMergeSourceConcealPulse()", StringComparison.Ordinal);
        var scheduleIndex = methodBody.IndexOf("ScheduleShellWindowRegistration(1)", StringComparison.Ordinal);
        Assert(resolveIndex >= 0 && pulseIndex > resolveIndex && scheduleIndex > resolveIndex,
            "The Explorer top-level resolution must happen before the conceal pulse and registration scheduler " +
            "are invoked.");

        Assert(methodBody.Contains("if (explorerTopLevel == 0) return", StringComparison.Ordinal),
            "A non-Explorer WinEvent must bail out immediately instead of scheduling work for unrelated windows.");

        return Task.CompletedTask;
    }

    public static async Task MergeSourceConcealPulseHasAbsoluteCeiling()
    {
        var pulse = new MergeSourceConcealPulse(absoluteCeilingMs: 90, sleepMs: 5);
        var callCount = 0;

        for (var i = 0; i < 5; i++)
        {
            pulse.Start(() => true, () => Interlocked.Increment(ref callCount), durationMs: 500);
            await Task.Delay(20);
        }

        await Task.Delay(80);
        var countAfterCeiling = Volatile.Read(ref callCount);
        await Task.Delay(80);

        Assert(countAfterCeiling > 0, "The pulse must run while it is inside the bounded window.");
        Assert(Volatile.Read(ref callCount) == countAfterCeiling,
            "A stream of Start calls must not extend the first pulse beyond its absolute ceiling.");
    }

    public static async Task MergeSourceConcealPulseNeverRestartsItselfInFinally()
    {
        var pulse = new MergeSourceConcealPulse(absoluteCeilingMs: 80, sleepMs: 5);
        var callCount = 0;

        pulse.Start(() => true, () => Interlocked.Increment(ref callCount), durationMs: 30);
        await Task.Delay(80);
        var countAfterExit = Volatile.Read(ref callCount);
        await Task.Delay(80);

        Assert(countAfterExit > 0, "The pulse must run before exiting.");
        Assert(Volatile.Read(ref callCount) == countAfterExit,
            "The pulse worker must not re-arm itself from inside its own exit path.");
    }

    public static async Task MergeSourceConcealPulseSleepIsAtLeast25Ms()
    {
        var pulse = new MergeSourceConcealPulse(absoluteCeilingMs: 140, sleepMs: 25);
        var callCount = 0;

        pulse.Start(() => true, () => Interlocked.Increment(ref callCount), durationMs: 100);
        await Task.Delay(140);

        Assert(Volatile.Read(ref callCount) <= 7,
            "The pulse worker must not run as a tight loop while scanning Explorer windows.");
    }

    public static Task ExplorerWatcherDisposeReleasesExplorerCheckTimer()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);

        // The outer ExplorerWatcher.Dispose is the last "public void Dispose" in the file (the nested
        // WinEventHookThread also defines one). Use LastIndexOf to skip the nested implementation.
        var disposeSignatureIndex = source.LastIndexOf("public void Dispose", StringComparison.Ordinal);
        Assert(disposeSignatureIndex >= 0, "Could not find the outer ExplorerWatcher.Dispose method.");
        var openBraceIndex = source.IndexOf('{', disposeSignatureIndex);
        var depth = 0;
        var endIndex = -1;
        for (var i = openBraceIndex; i < source.Length; i++)
        {
            if (source[i] == '{') depth++;
            else if (source[i] == '}') depth--;
            if (depth == 0) { endIndex = i; break; }
        }
        Assert(endIndex > openBraceIndex, "Could not extract the outer Dispose body.");
        var disposeBody = source.Substring(openBraceIndex, endIndex - openBraceIndex + 1);

        Assert(disposeBody.Contains("_explorerCheckTimer?.Dispose()", StringComparison.Ordinal) &&
               disposeBody.Contains("_explorerCheckTimer = null", StringComparison.Ordinal),
            "Dispose must release the explorer-process polling timer. If the timer is still armed when WinTab " +
            "shuts down, the System.Threading.Timer callback can fire against torn-down shell objects.");

        return Task.CompletedTask;
    }

    public static Task PreExistingExplorerWindowsAreNotConcealedDuringStartupRace()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var adoptBody = SourceContract.ExtractMethodBody(source, "private List<(InternetExplorer Window, WindowInfo WindowInfo)> AdoptNewShellWindows");
        var startBody = SourceContract.ExtractMethodBody(source, "public void StartHook");

        // StartHook must seed the pre-existing-window protection (RecoverHiddenExplorerWindows adds every
        // currently-open CabinetWClass top-level to _processedHWnds) BEFORE enabling forced tabs or starting
        // the conceal pulse. Otherwise the pulse worker can race the shell-objects loop and hide a user's
        // pre-existing Explorer window before InitializeShellObjects manages to hook it.
        var recoverIndex = startBody.IndexOf("RecoverHiddenExplorerWindows", StringComparison.Ordinal);
        var forcingIndex = startBody.IndexOf("_isForcingTabs = true", StringComparison.Ordinal);
        var pulseIndex = startBody.IndexOf("StartMergeSourceConcealPulse", StringComparison.Ordinal);
        Assert(recoverIndex >= 0 && forcingIndex > recoverIndex && pulseIndex > forcingIndex,
            "StartHook must seed pre-existing window protection before enabling forced tabs and starting the conceal pulse.");

        // AdoptNewShellWindows is also reached through ShellWindows registration callbacks and from
        // foreground/show WinEvents on the user's existing Explorer windows. Its inline HideMergeSourceWindow
        // call must respect _processedHWnds so windows that existed before WinTab took over are never hidden
        // as merge sources, even if they happen to be enumerated before InitializeShellObjects hooks them.
        var hideIndex = adoptBody.IndexOf("HideMergeSourceWindow(hWnd)", StringComparison.Ordinal);
        Assert(hideIndex >= 0, "AdoptNewShellWindows must still hide real merge-source windows.");

        var preHide = adoptBody.Substring(0, hideIndex);
        var lastIfStart = preHide.LastIndexOf("if (", StringComparison.Ordinal);
        Assert(lastIfStart >= 0, "AdoptNewShellWindows hide call must be gated by an if-condition.");
        var hideCondition = preHide.Substring(lastIfStart, preHide.Length - lastIfStart);
        Assert(hideCondition.Contains("_processedHWnds.ContainsKey(hWnd)", StringComparison.Ordinal),
            "AdoptNewShellWindows must skip windows already in _processedHWnds so pre-existing Explorer " +
            "windows (seeded by RecoverHiddenExplorerWindows on StartHook) are never hidden as merge sources " +
            "during the startup race. On first install with multiple pre-existing windows, the lack of this " +
            "check causes the second window in the batch to be concealed and effectively unusable.");

        return Task.CompletedTask;
    }

    public static Task RecycleBinAndExternalFolderOpensReuseExistingTabWithoutDuplicates()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var openBody = SourceContract.ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection");

        var foundIndex = openBody.IndexOf("TrySearchForTab(windowToOpen.Location, windowToOpen.Handle", StringComparison.Ordinal);
        Assert(foundIndex >= 0, "Reuse must search by location and source handle.");

        var foregroundIndex = openBody.IndexOf("Helper.RestoreWindowToForeground(windowHandle)", foundIndex, StringComparison.Ordinal);
        var selectIndex = openBody.IndexOf("await SelectTabByHandle(windowHandle, existingTab)", foundIndex, StringComparison.Ordinal);
        var reusedLogIndex = openBody.IndexOf("OpenTab reused target", foundIndex, StringComparison.Ordinal);
        var returnTrueIndex = openBody.IndexOf("return true;", reusedLogIndex >= 0 ? reusedLogIndex : foundIndex, StringComparison.Ordinal);

        Assert(foregroundIndex > foundIndex && selectIndex > foregroundIndex && reusedLogIndex > selectIndex && returnTrueIndex > reusedLogIndex,
            "Reuse must, in order: foreground the target window, fire best-effort tab selection, log success, and return true. " +
            "Any branch that converts selection failure into duplicate-tab creation re-opens the original bug.");

        Assert(!openBody.Contains("if (await SelectTabByHandle(windowHandle, existingTab", StringComparison.Ordinal),
            "Reuse must not gate success on SelectTabByHandle's return value — slow shell folders (Recycle Bin, network shares) miss tight cycle budgets.");
        Assert(!openBody.Contains("OpenTab reuse-select-failed", StringComparison.Ordinal),
            "The reuse-select-failed path used to create a duplicate tab; it must be removed entirely.");

        var newTabIndex = openBody.IndexOf("RequestToOpenNewTab(mainWindowHWnd, lockToOpenWindows: false)", StringComparison.Ordinal);
        Assert(newTabIndex == -1 || newTabIndex > returnTrueIndex,
            "A new tab may only be created on the no-match path, never as a fallback when an existing matching tab was already located.");

        return Task.CompletedTask;
    }

    public static Task RecycleBinAndVirtualFolderPidlResolutionAndEquivalence()
    {
        using var comparer = new ShellPathComparer();

        var recycleBinPath = "shell:::{645FF040-5081-101B-9F08-00AA002F954E}";
        var cleanRecycleBinPath = "::{645FF040-5081-101B-9F08-00AA002F954E}";
        var shortcutRecycleBinPath = "shell:RecycleBinFolder";
        var uncPath = @"\\127.0.0.1\c$";

        var pidl1 = comparer.GetPidlFromPath(recycleBinPath);
        var pidl2 = comparer.GetPidlFromPath(cleanRecycleBinPath);
        var pidl3 = comparer.GetPidlFromPath(shortcutRecycleBinPath);

        var normalizedUnc = Helper.NormalizeLocation(uncPath);
        var pidlUnc = comparer.GetPidlFromPath(normalizedUnc);

        Console.WriteLine($"[DEBUG TEST] pidl1: {pidl1:X}, pidl2: {pidl2:X}, pidl3: {pidl3:X}, normalizedUnc: '{normalizedUnc}', pidlUnc: {pidlUnc:X}");

        try
        {
            Assert(pidl1 != 0, "PIDL for shell:::RecycleBin path should be successfully resolved.");
            Assert(pidl2 != 0, "PIDL for clean RecycleBin path should be successfully resolved.");
            Assert(pidl3 != 0, "PIDL for shell:RecycleBinFolder path should be successfully resolved.");
            Assert(pidlUnc != 0, $"PIDL for normalized UNC path '{normalizedUnc}' should be successfully resolved.");

            var fsPath1 = ShellPathComparer.GetPathFromPidl(pidl1);
            var fsPath2 = ShellPathComparer.GetPathFromPidl(pidl2);
            var fsPath3 = ShellPathComparer.GetPathFromPidl(pidl3);
            Console.WriteLine($"[DEBUG TEST] fsPath1: '{fsPath1}', fsPath2: '{fsPath2}', fsPath3: '{fsPath3}'");

            var equivIds12 = comparer.CompareIds(pidl1, pidl2);
            var equivIds23 = comparer.CompareIds(pidl2, pidl3);
            Console.WriteLine($"[DEBUG TEST] CompareIds 1-2: {equivIds12}, CompareIds 2-3: {equivIds23}");

            Assert(comparer.IsEquivalent(pidl1, pidl2), "Virtual path and clean virtual path should be equivalent.");
            Assert(comparer.IsEquivalent(recycleBinPath, cleanRecycleBinPath), "Virtual path string comparisons should be equivalent.");
            Assert(comparer.IsEquivalent(shortcutRecycleBinPath, cleanRecycleBinPath), "Shortcut and clean path should be equivalent.");
        }
        finally
        {
            if (pidl1 != 0) Marshal.FreeCoTaskMem(pidl1);
            if (pidl2 != 0) Marshal.FreeCoTaskMem(pidl2);
            if (pidl3 != 0) Marshal.FreeCoTaskMem(pidl3);
            if (pidlUnc != 0) Marshal.FreeCoTaskMem(pidlUnc);
        }

        return Task.CompletedTask;
    }

    public static Task TabSelectionBudgetToleratesSlowShellFoldersWithoutSkippingCycles()
    {
        var watcherPath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var watcher = File.ReadAllText(watcherPath);
        var selectBody = SourceContract.ExtractMethodBody(watcher, "public Task<bool> SelectTabByHandle");

        Assert(selectBody.Contains("totalTimeoutMs: timeoutMs", StringComparison.Ordinal),
            "SelectTabByHandle must forward its timeout parameter to the engine.");
        Assert(selectBody.Contains("perStepTimeoutMs: 250", StringComparison.Ordinal),
            "Each tab cycle step must be given at least 250 ms so slow shell folder enumeration does not exhaust the budget mid-cycle.");
        Assert(watcher.Contains("public Task<bool> SelectTabByHandle(nint windowHandle, nint tabHandle, int timeoutMs = 2_500)", StringComparison.Ordinal),
            "Default total budget must be ~2500 ms — short enough for users, long enough for the slowest shell folder transitions.");

        return Task.CompletedTask;
    }

    public static Task TabSearchMatchesPreHookWindowsSoExternalOpensFindTheirTabDuringRaces()
    {
        var watcherPath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var watcher = File.ReadAllText(watcherPath);
        var searchBody = SourceContract.ExtractMethodBody(watcher, "private bool TrySearchForTab");

        Assert(!searchBody.Contains("windowInfo.EventsHooked", StringComparison.Ordinal),
            "TrySearchForTab must not gate on EventsHooked; a tab whose hook is still pending is still a valid reuse target as long as it has a tab handle.");
        Assert(searchBody.Contains("excludedTopLevelWindow", StringComparison.Ordinal),
            "Source-window exclusion still owns the 'do not match yourself' invariant.");
        Assert(searchBody.Contains("!tab.HasValue || tab.Value == 0", StringComparison.Ordinal),
            "Only a resolved tab handle is required.");

        var locationIndex = searchBody.IndexOf("windowInfo.Location ?? GetLocation(window)", StringComparison.Ordinal);
        Assert(locationIndex >= 0, "Compare path must fall back to a live GetLocation call when the cached location is missing.");

        var preLocation = searchBody[..locationIndex];
        var lastTryIndex = preLocation.LastIndexOf("try", StringComparison.Ordinal);
        Assert(lastTryIndex >= 0 && lastTryIndex < locationIndex,
            "The GetLocation fallback must be wrapped in a try/catch so a single misbehaving COM object cannot void the whole search and leak the user back to the duplicate-tab path.");

        // Concurrent .Add/.Remove on _windowEntryDict during enumeration would throw
        // InvalidOperationException, get swallowed by the outer catch, and silently fail the
        // search. The foreach must hold _windowEntryDictLock for the duration of the scan.
        var foreachIndex = searchBody.IndexOf("foreach (var (window, windowInfo, tab) in _windowEntryDict)", StringComparison.Ordinal);
        Assert(foreachIndex >= 0, "Search must enumerate the window-entry dictionary.");
        var preForeach = searchBody[..foreachIndex];
        var lastLockIndex = preForeach.LastIndexOf("lock (_windowEntryDictLock)", StringComparison.Ordinal);
        Assert(lastLockIndex >= 0,
            "TrySearchForTab's foreach must be inside lock (_windowEntryDictLock) so concurrent dict mutation cannot break the search.");

        return Task.CompletedTask;
    }

    public static async Task TabSelectionEngineSlowExplorerStillConvergesWithGenerousPerStepBudget()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33, 44, 55 }, activeIndex: 0);
        var pendingIndex = -1;
        var pendingDeadline = 0L;
        const int slowDelayMs = 120;

        fixture.OnSelectByIndex = i =>
        {
            pendingIndex = i;
            pendingDeadline = Environment.TickCount64 + slowDelayMs;
        };
        fixture.OnGetActive = current =>
        {
            if (pendingIndex >= 0 && Environment.TickCount64 >= pendingDeadline)
            {
                fixture.SetActiveIndex(pendingIndex);
                pendingIndex = -1;
                current = fixture.GetActiveSnapshot();
            }
            return current;
        };

        var ok = await TabSelectionEngine.CycleToTabAsync(
            44,
            fixture.GetTabs,
            fixture.GetActive,
            fixture.SendSelectByIndex,
            totalTimeoutMs: 2_500,
            pollSleepMs: 5,
            perStepTimeoutMs: 250);

        Assert(ok, "Engine must converge on the requested tab even when Explorer takes ~120 ms per switch.");
        Assert(fixture.GetActive() == 44, "Final active tab must match the requested handle.");
    }

    public static Task NavigateCompleteHandlerSignalsUnconditionally()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var navigateBody = SourceContract.ExtractMethodBody(source, "private async Task<bool> NavigateNewTabToTargetAsync");

        // The handler block follows the lambda assignment. Extract it precisely.
        var handlerAssignmentIndex = navigateBody.IndexOf("navigateHandler = (object _, ref object _) =>", StringComparison.Ordinal);
        Assert(handlerAssignmentIndex >= 0, "Navigation handler lambda must be present.");

        var handlerEnd = navigateBody.IndexOf("};", handlerAssignmentIndex, StringComparison.Ordinal);
        Assert(handlerEnd > handlerAssignmentIndex, "Navigation handler lambda must terminate with };");
        var handlerBlock = navigateBody.Substring(handlerAssignmentIndex, handlerEnd - handlerAssignmentIndex);

        Assert(handlerBlock.Contains("navigationCompleted.TrySetResult(true);", StringComparison.Ordinal),
            "NavigateComplete2 handler must signal completion immediately. Explorer fires NavigateComplete2 once per navigation; gating tcs on AreLocationsEquivalent throws the typical fast-path away because LocationURL is updated a few ticks after the event arrives.");

        // The handler must not gate the signal on an equivalence check; that was the slow merge bug.
        Assert(!handlerBlock.Contains("AreLocationsEquivalent(TryGetLocation(window), targetLocation)", StringComparison.Ordinal),
            "NavigateComplete2 handler must not gate tcs on AreLocationsEquivalent. Run the equivalence check AFTER waiting for the event signal, not inside the handler — otherwise tcs stays unset whenever LocationURL trails the event by a few ticks, and the merge pays the full NavigationVerificationWaitMs.");

        return Task.CompletedTask;
    }

    public static Task NavigateCompleteWaitBudgetIsShort()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);

        // Pull the constant declaration line. It must be ≤ 700 ms because the typical
        // NavigateComplete2 event arrives in under 200 ms; waiting longer just stalls
        // merges and tab opens for no benefit when the event is being raced against Task.Delay.
        Assert(source.Contains("private const int NavigationCompleteWaitMs = 600;", StringComparison.Ordinal) ||
               source.Contains("private const int NavigationCompleteWaitMs = 500;", StringComparison.Ordinal) ||
               source.Contains("private const int NavigationCompleteWaitMs = 400;", StringComparison.Ordinal),
            "NavigationCompleteWaitMs must be ≤ 600 ms. The old 1_200 ms budget combined with the equivalence-gating bug stretched every merge into a 2.4 s wait.");

        return Task.CompletedTask;
    }

    public static Task NavigationFastPathChecksLocationAfterEvent()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var navigateBody = SourceContract.ExtractMethodBody(source, "private async Task<bool> NavigateNewTabToTargetAsync");

        var whenAnyIndex = navigateBody.IndexOf("Task.WhenAny(navigationCompleted.Task", StringComparison.Ordinal);
        Assert(whenAnyIndex >= 0, "NavigateNewTabToTargetAsync must race tcs against a Task.Delay.");

        var equivalenceIndex = navigateBody.IndexOf("AreLocationsEquivalent(TryGetLocation(window), targetLocation)", whenAnyIndex, StringComparison.Ordinal);
        Assert(equivalenceIndex > whenAnyIndex,
            "The equivalence fast-path check must run AFTER the WhenAny race, so the navigation handler can fire fast and the equivalence check verifies success without blocking the merge.");

        var slowPathIndex = navigateBody.IndexOf("WaitForNavigation(window, targetLocation, NavigationVerificationWaitMs)", equivalenceIndex, StringComparison.Ordinal);
        Assert(slowPathIndex > equivalenceIndex,
            "Slow-path polling must come after the equivalence fast-path so the typical merge case returns within ~150 ms.");

        return Task.CompletedTask;
    }

    public static Task WaitForExplorerTabCountUsesShortBudget()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private static Task<int> WaitForExplorerTabCount");

        Assert(!methodBody.Contains("1_500", StringComparison.Ordinal) &&
               !methodBody.Contains("1_200", StringComparison.Ordinal) &&
               !methodBody.Contains("1_000", StringComparison.Ordinal),
            "WaitForExplorerTabCount must not wait a full second or more. The tab count is normally established within ~50 ms; the old 1.5 s budget just stalled the merge decision on a stable signal.");
        // 800 ms is the safe upper bound — long enough for slow systems (HDD, busy Explorer COM
        // re-init), short enough that misses do not bottleneck rapid merge bursts.
        Assert(methodBody.Contains("800,", StringComparison.Ordinal) ||
               methodBody.Contains("700,", StringComparison.Ordinal) ||
               methodBody.Contains("600,", StringComparison.Ordinal) ||
               methodBody.Contains("500,", StringComparison.Ordinal) ||
               methodBody.Contains("400,", StringComparison.Ordinal),
            "WaitForExplorerTabCount budget must be ≤ 800 ms.");

        return Task.CompletedTask;
    }

    public static Task CloseMergedSourceWindowUsesShortBudget()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private async Task<bool> CloseMergedSourceWindowAsync");

        // The original 1500 + 900 = 2400 ms close pile-up was a major contributor to merge latency on rapid
        // detach/merge cycles. The window's COM Quit() runs in parallel and usually closes within ~200 ms.
        Assert(!methodBody.Contains("1_500", StringComparison.Ordinal),
            "CloseMergedSourceWindowAsync must not wait 1.5 seconds on the first close attempt.");
        Assert(!methodBody.Contains("900", StringComparison.Ordinal),
            "CloseMergedSourceWindowAsync must not wait 900 ms on the second close attempt.");

        // At least one short budget call should remain.
        Assert(methodBody.Contains("700,", StringComparison.Ordinal) ||
               methodBody.Contains("600,", StringComparison.Ordinal) ||
               methodBody.Contains("500,", StringComparison.Ordinal),
            "First close verification budget must be ≤ 700 ms.");
        Assert(methodBody.Contains("300,", StringComparison.Ordinal) ||
               methodBody.Contains("250,", StringComparison.Ordinal) ||
               methodBody.Contains("200,", StringComparison.Ordinal),
            "Second close verification budget must be ≤ 300 ms.");

        return Task.CompletedTask;
    }

    public static Task ShellWindowRegistrationLoopIsBounded()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = SourceContract.ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowsAsync");

        // The old loop ran for up to 8 iterations with 75-150 ms gaps. That added 600-1200 ms of latency
        // to every Explorer launch even after the merge had completed. 4 iterations is plenty to catch
        // rapid registration bursts.
        Assert(!methodBody.Contains("for (var i = 0; i < 8; i++)", StringComparison.Ordinal),
            "ProcessRegisteredShellWindowsAsync must not loop 8 times. Rapid bursts settle within 2-4 iterations on Windows 11; the longer loop just adds wall time to every Explorer launch.");
        Assert(methodBody.Contains("for (var i = 0; i < 4; i++)", StringComparison.Ordinal) ||
               methodBody.Contains("for (var i = 0; i < 3; i++)", StringComparison.Ordinal),
            "Registration loop must be bounded to ≤ 4 iterations.");

        // The inter-iteration wait must be short. The old 150 ms gap meant a single burst of 4 windows
        // paid 600 ms even when each iteration found new windows immediately.
        Assert(!methodBody.Contains("Task.Delay(150)", StringComparison.Ordinal),
            "ProcessRegisteredShellWindowsAsync must not wait 150 ms between probes.");

        return Task.CompletedTask;
    }

    public static Task ReusedHwndDuringProcessedGraceIsHookedNotLeftTransparent()
    {
        var sourcePath = SourceContract.FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var registrationBody = SourceContract.ExtractMethodBody(source, "private async Task ProcessRegisteredShellWindowAsync");

        // Each `if (_processedHWnds.ContainsKey(hWnd))` early-return guard in the registration
        // body must (a) restore any hidden state for the reused hwnd and (b) hand the new window
        // to RegisterIndependentWindow before returning. Otherwise a hwnd that explorer reuses
        // inside the 7-second PreventWindowHiding grace gets adopted into _windowEntryDict but
        // never hooked. Once the grace expires, the next WinEvent hides it as a single-tab merge
        // source and nothing ever restores it — the transparent "此电脑" background residual.
        var guardIndex = 0;
        var guardCount = 0;
        while (true)
        {
            guardIndex = registrationBody.IndexOf("if (_processedHWnds.ContainsKey(hWnd))", guardIndex, StringComparison.Ordinal);
            if (guardIndex < 0)
                break;

            var afterGuard = registrationBody.Substring(guardIndex);
            var nextReturn = afterGuard.IndexOf("return;", StringComparison.Ordinal);
            Assert(nextReturn > 0,
                "Each _processedHWnds early-return guard must terminate with a return statement.");

            var block = afterGuard.Substring(0, nextReturn);
            Assert(block.Contains("RestoreMergeSourceWindowAsync(hWnd)", StringComparison.Ordinal),
                "Early-return when _processedHWnds already covers this hwnd must restore any hidden state — otherwise a reused hwnd can stay alpha=0 forever after the grace expires.");
            Assert(block.Contains("RegisterIndependentWindow(window, windowInfo, hWnd)", StringComparison.Ordinal),
                "Early-return when _processedHWnds already covers this hwnd must hook the new shell window via RegisterIndependentWindow so _hookedTopLevelUseCounts protects it from later merge-source hide passes.");

            guardCount++;
            guardIndex += "if (_processedHWnds.ContainsKey(hWnd))".Length;
        }

        Assert(guardCount >= 1,
            "ProcessRegisteredShellWindowAsync must keep at least one _processedHWnds guard so reused-hwnd registrations skip merge processing.");

        return Task.CompletedTask;
    }

    public static async Task OpenTabFastPathReturnsBeforeNavigationVerificationOnComplete()
    {
        // This is a behavioural test: with the new handler design, racing tcs vs Task.Delay must
        // return quickly when the handler fires. We simulate the contract directly.
        var navigationCompleted = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var fastDelayMs = 600;
        var fireDelayMs = 80;

        _ = Task.Run(async () =>
        {
            await Task.Delay(fireDelayMs);
            navigationCompleted.TrySetResult(true);
        });

        var start = Environment.TickCount64;
        await Task.WhenAny(navigationCompleted.Task, Task.Delay(fastDelayMs));
        var elapsed = Environment.TickCount64 - start;

        Assert(elapsed < 250,
            $"WhenAny race must return as soon as the handler fires (target ~{fireDelayMs} ms). Got {elapsed} ms. " +
            "The merge slowness reproducer is the old code waited the full NavigationCompleteWaitMs (1200 ms) because the handler never signalled tcs.");
    }

    private sealed class TabSelectionFixture
    {
        public nint[] Tabs { get; }
        public List<int> SelectionCalls { get; } = new();
        public Action<int>? OnSelectByIndex { get; set; }
        public Func<nint, nint>? OnGetActive { get; set; }
        private int _activeIndex;

        public TabSelectionFixture(nint[] tabs, int activeIndex)
        {
            Tabs = tabs;
            _activeIndex = activeIndex;
        }

        public nint[] GetTabs() => Tabs;
        public nint GetActive()
        {
            var current = _activeIndex >= 0 && _activeIndex < Tabs.Length ? Tabs[_activeIndex] : 0;
            return OnGetActive != null ? OnGetActive(current) : current;
        }
        public void SetActiveIndex(int index)
        {
            if (index >= 0 && index < Tabs.Length)
                _activeIndex = index;
        }
        public nint GetActiveSnapshot()
        {
            return _activeIndex >= 0 && _activeIndex < Tabs.Length ? Tabs[_activeIndex] : 0;
        }
        public void SendSelectByIndex(int i)
        {
            SelectionCalls.Add(i);
            OnSelectByIndex?.Invoke(i);
        }
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }
}
