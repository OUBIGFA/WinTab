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

return await ExplorerLaunchLocationResolverTests.RunAll();

internal static class ExplorerLaunchLocationResolverTests
{
    public static async Task<int> RunAll()
    {
        var tests = new (string Name, Func<Task> Body)[]
        {
            ("waits for the real folder when Explorer first reports This PC", WaitsForRealFolderAfterTransientDefault),
            ("returns the default folder only after the startup location stays default", ReturnsDefaultAfterTimeout),
            ("waits for a non-default location to stabilize", WaitsForStableNonDefaultLocation),
            ("normalizes file URLs to local filesystem paths", NormalizesFileUrlsToLocalPaths),
            ("keeps web URLs usable", KeepsWebUrlsUsable),
            ("SelectTabByHandle does not activate intermediate tabs", ExplorerTabSelectionTests.SelectTabByHandleDoesNotActivateIntermediateTabs),
            ("tab merge uses the ExplorerTabUtility fast path", ExplorerTabSelectionTests.TabMergeUsesExplorerTabUtilityFastPath),
            ("failed navigation closes the transient This PC tab", ExplorerTabSelectionTests.FailedNavigationClosesTransientThisPcTab),
            ("ExplorerTabUtility-style show hook does not hide independent windows", ExplorerTabSelectionTests.ShowHookDoesNotHideIndependentWindows),
            ("ShellWindows registration callback stays non-blocking", ExplorerTabSelectionTests.ShellWindowRegistrationCallbackStaysNonBlocking),
            ("default-location Explorer windows are released instead of hidden indefinitely", ExplorerTabSelectionTests.DefaultLocationWindowsAreReleasedInsteadOfHiddenIndefinitely),
            ("HideWindow reapplies transparency after Explorer resets styles", ExplorerTabSelectionTests.HideWindowReappliesTransparency),
            ("hidden Explorer windows are restored on lifecycle boundaries", ExplorerTabSelectionTests.HiddenExplorerWindowsAreRestoredOnLifecycleBoundaries),
            ("orphaned transparent Explorer windows are recovered without cache", ExplorerTabSelectionTests.OrphanedTransparentExplorerWindowsAreRecoveredWithoutCache),
            ("tab reuse excludes only the merge source instead of young tabs", ExplorerTabSelectionTests.TabReuseExcludesMergeSourceInsteadOfYoungTabs),
            ("tab reuse never creates a duplicate after finding an existing path", ExplorerTabSelectionTests.TabReuseDoesNotDuplicateAfterSelectionFailure)
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
            PollIntervalMs: 1));
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

    private static class KnownLocations
    {
        public const string ThisPc = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
        public const string Downloads = "file:///C:/Users/BIGFA/Downloads";
        public const string TargetFolder = "file:///E:/WinTabStress/Target";
    }
}

internal static class ExplorerTabSelectionTests
{
    public static Task SelectTabByHandleDoesNotActivateIntermediateTabs()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "public async Task SelectTabByHandle") +
                         ExtractMethodBody(source, "private async Task<bool> TrySelectTabByHandleDirectAsync");

        Assert(!methodBody.Contains("SelectTabByIndex(windowHandle, i)", StringComparison.Ordinal),
            "SelectTabByHandle must not select every tab index while searching for the target.");
        Assert(!methodBody.Contains("i < tabs.Length", StringComparison.Ordinal),
            "SelectTabByHandle must not linearly scan tab indexes; reuse must activate the target directly.");
        Assert(!methodBody.Contains("SelectTabByIndex(windowHandle, tabIndex)", StringComparison.Ordinal),
            "SelectTabByHandle must not guess a visual tab from ShellTabWindowClass z-order.");
        Assert(!methodBody.Contains("while", StringComparison.Ordinal),
            "SelectTabByHandle must not retry in a loop; a failed direct activation should fail fast.");

        return Task.CompletedTask;
    }

    public static Task TabMergeUsesExplorerTabUtilityFastPath()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection") +
                         ExtractMethodBody(source, "private async Task<bool> NavigateNewTabToTargetAsync");

        Assert(methodBody.Contains("SelectTabByUniqueNameVerified(windowHandle, existingTab, 500, existingWindow)", StringComparison.Ordinal),
            "Tab reuse should follow ExplorerTabUtility's direct handle activation path.");
        Assert(methodBody.Contains("ListenForNewExplorerTabAsync(mainWindowHWnd, currentTabs, 2_000)", StringComparison.Ordinal),
            "New tab creation should use the short ExplorerTabUtility wait window.");
        Assert(!methodBody.Contains("restore-previous", StringComparison.OrdinalIgnoreCase),
            "The merge path must not restore previous tabs while opening a target tab.");
        Assert(!methodBody.Contains("CloseUnexpectedDefaultTabs", StringComparison.Ordinal),
            "The merge path must not run a broad default-tab cleanup pass.");
        Assert(!methodBody.Contains("newTabTimeoutMs", StringComparison.Ordinal),
            "The merge path must not carry the old long timeout path.");
        Assert(!source.Contains("SelectTabByKnownIndexVerified", StringComparison.Ordinal),
            "The known-index merge path should be removed, not bypassed.");

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
                         ExtractMethodBody(source, "private async Task<bool> TryAutoMergePendingWindow");

        Assert(methodBody.Contains("WaitForNavigation(window, targetLocation, 5_000)", StringComparison.Ordinal) &&
               methodBody.Contains("AreLocationsEquivalent(TryGetLocation(window), targetLocation)", StringComparison.Ordinal),
            "The newly opened tab must be checked against the requested target folder.");
        Assert(methodBody.Contains("CloseFailedNewTabAsync(mainWindowHWnd, newTabHandle)", StringComparison.Ordinal),
            "A tab that remains at This PC or another wrong location must be closed instead of left behind.");
        Assert(methodBody.Contains("Helper.IsFileExplorerForeground(out var foregroundWindow)", StringComparison.Ordinal),
            "The merge target must prefer the current foreground Explorer instead of a stale background This PC window.");
        Assert(methodBody.Contains("IsStableMergeTargetWindow(foregroundWindow, otherThan)", StringComparison.Ordinal),
            "The foreground window must be a stable merge target, not another transient source window.");
        Assert(methodBody.Contains("GetMainWindowHWnd(hWnd)", StringComparison.Ordinal),
            "Auto-merge must resolve the target Explorer while excluding the transient source window.");
        Assert(methodBody.Contains("IsFallbackMergeTargetWindow(h, otherThan)", StringComparison.Ordinal),
            "If the strict cached target is unavailable, merge should fall back to ExplorerTabUtility's tagged-window selection instead of opening a new window.");
        Assert(methodBody.Contains("Helper.NormalizeLocation(location)", StringComparison.Ordinal) &&
               methodBody.Contains("Helper.NormalizeLocation(targetLocation)", StringComparison.Ordinal),
            "Navigation completion checks must normalize local paths and file URLs before comparing them.");

        return Task.CompletedTask;
    }

    public static Task ShowHookDoesNotHideIndependentWindows()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "private bool TryHideIncomingExplorerWindow") +
                         ExtractMethodBody(source, "private bool IsRegisteredIndependentExplorerWindow") +
                         ExtractMethodBody(source, "private void PreventWindowHiding") +
                         ExtractMethodBody(source, "private void OnWindowShown") +
                         ExtractMethodBody(source, "private void InitializeShellObjects") +
                         ExtractMethodBody(source, "private sealed class WinEventHookThread");

        Assert(methodBody.Contains("IsRegisteredIndependentExplorerWindow(hWnd)", StringComparison.Ordinal),
            "The early conceal path must not hide an Explorer window that was already registered as independent.");
        Assert(methodBody.Contains("_registeredIndependentHWnds.ContainsKey(hWnd)", StringComparison.Ordinal) &&
               methodBody.Contains("_registeredIndependentHWnds[hWnd] = 0", StringComparison.Ordinal),
            "Independent Explorer windows should be protected by an explicit handle cache, not a slow COM scan.");
        Assert(methodBody.Contains("SetWinEventHook(WinApi.EVENT_OBJECT_SHOW", StringComparison.Ordinal) &&
               methodBody.Contains("SetWinEventHook(WinApi.EVENT_OBJECT_CREATE", StringComparison.Ordinal),
            "Window concealment should stay event-driven instead of polling Explorer windows.");
        Assert(!source.Contains("RunConcealWatchdog", StringComparison.Ordinal) &&
               !source.Contains("RunConcealSweepAsync", StringComparison.Ordinal),
            "The merge path should not depend on custom high-frequency window sweeps.");
        Assert(methodBody.Contains("Priority = ThreadPriority.Highest", StringComparison.Ordinal),
            "The WinEvent hook thread should run at high priority so conceal events are handled before Explorer paints.");

        return Task.CompletedTask;
    }

    public static Task ShellWindowRegistrationCallbackStaysNonBlocking()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "private void OnShellWindowRegistered");

        Assert(methodBody.Contains("SchedulePendingAutoMerge(1)", StringComparison.Ordinal),
            "ShellWindows registration should schedule merge work after returning to Explorer.");
        Assert(!methodBody.Contains("AdoptNewShellWindowsForImmediateConceal", StringComparison.Ordinal),
            "ShellWindows registration must not synchronously enumerate ShellWindows from Explorer's COM event callback.");

        return Task.CompletedTask;
    }

    public static Task DefaultLocationWindowsAreReleasedInsteadOfHiddenIndefinitely()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var methodBody = ExtractMethodBody(source, "private async Task<bool> TryAutoMergePendingWindow");
        var hideBody = ExtractMethodBody(source, "private bool TryHideIncomingExplorerWindow");

        var firstStartupIndex = methodBody.IndexOf("IsStartupExplorerLocation(firstLocation)", StringComparison.Ordinal);
        var graceIndex = methodBody.IndexOf("PendingStartupLocationGraceMs", firstStartupIndex >= 0 ? firstStartupIndex : 0, StringComparison.Ordinal);
        var releaseIndex = methodBody.IndexOf("TryShowAsIndependentWindow(window, windowInfo)", graceIndex >= 0 ? graceIndex : 0, StringComparison.Ordinal);
        var resolveIndex = methodBody.IndexOf("ResolveInitialLocation(window)", StringComparison.Ordinal);
        Assert(firstStartupIndex >= 0 && graceIndex > firstStartupIndex && releaseIndex > graceIndex,
            "A pending Explorer window that stays at the default location must be released after a short startup grace window.");
        Assert(resolveIndex > releaseIndex,
            "Default-location waits must not block the global merge queue while other new Explorer windows need concealment.");
        Assert(!methodBody.Contains("180_000", StringComparison.Ordinal),
            "Hidden merge candidates must not stay concealed for minutes.");
        Assert(hideBody.Contains("HasTrackedIndependentTopLevelWindow(hWnd)", StringComparison.Ordinal),
            "The early WinEvent hide path must not conceal already-tracked stable Explorer windows.");

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
        var showIndependentBody = ExtractMethodBody(source, "private void TryShowAsIndependentWindow");
        var removeBody = ExtractMethodBody(source, "private void RemoveWindowAndUnhookEvents");

        var startupRestoreIndex = initializeBody.IndexOf("RecoverHiddenExplorerWindows", StringComparison.Ordinal);
        var eventHookIndex = initializeBody.IndexOf("WindowRegistered +=", StringComparison.Ordinal);
        Assert(startupRestoreIndex >= 0 && eventHookIndex >= 0 && startupRestoreIndex < eventHookIndex,
            "Startup must recover orphaned hidden Explorer windows before registering new merge hooks.");
        Assert(stopBody.Contains("RecoverHiddenExplorerWindows", StringComparison.Ordinal),
            "StopHook must release hidden Explorer windows before disabling forced tabs.");
        Assert(disposeBody.Contains("RecoverHiddenExplorerWindows", StringComparison.Ordinal),
            "DisposeShellObjects must release hidden Explorer windows before COM objects are dropped.");
        Assert(showIndependentBody.Contains("RestoreHiddenExplorerWindowAsync", StringComparison.Ordinal),
            "Windows that stop being merge candidates must be restored through the guarded restore path.");
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

    public static Task TabReuseExcludesMergeSourceInsteadOfYoungTabs()
    {
        var sourcePath = FindRepoFile("WinTab", "Hooks", "ExplorerWatcher.cs");
        var source = File.ReadAllText(sourcePath);
        var searchBody = ExtractMethodBody(source, "private bool TrySearchForTab");
        var openBody = ExtractMethodBody(source, "private async Task<bool> OpenTabNavigateWithSelection");

        Assert(!searchBody.Contains("CreatedAt", StringComparison.Ordinal),
            "Tab reuse must not exclude freshly-created real target tabs by age.");
        Assert(searchBody.Contains("windowInfo.CanAutoMerge", StringComparison.Ordinal),
            "Tab reuse should exclude pending merge-source windows by state.");
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
        var selectBody = ExtractMethodBody(source, "private async Task<bool> SelectTabByUniqueNameVerified") +
                         ExtractMethodBody(source, "private async Task<bool> TrySelectTabByKnownOrderAsync");

        var failureLogIndex = openBody.IndexOf("OpenTab reuse-select-failed", StringComparison.Ordinal);
        var duplicateOpenIndex = openBody.IndexOf("RequestToOpenNewTab", StringComparison.Ordinal);
        var returnTrueAfterFailureIndex = openBody.IndexOf("return true;", failureLogIndex >= 0 ? failureLogIndex : 0, StringComparison.Ordinal);
        Assert(failureLogIndex >= 0 && returnTrueAfterFailureIndex > failureLogIndex && returnTrueAfterFailureIndex < duplicateOpenIndex,
            "Once reuse finds an existing path, selection failure must not fall through to opening a duplicate tab.");
        Assert(selectBody.Contains("GetActiveTabHandle(windowHandle) == tabHandle", StringComparison.Ordinal),
            "Reuse selection should succeed immediately when the target tab is already active.");
        Assert(selectBody.Contains("TrySelectTabByKnownOrderAsync", StringComparison.Ordinal),
            "Duplicate tab titles need a non-title fallback that still avoids tab-by-tab cycling.");

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

        using var app = StartWinTab(appPath, debugLog);
        try
        {
            await Task.Delay(startupDelayMs);

            StartExplorer(target);
            var firstTargetWindow = await WaitForTargetCountAsync(target, expectedCount: 1, timeoutMs: 12_000);
            if (firstTargetWindow.Count(window => IsSameFolder(window, target)) != 1)
            {
                Console.Error.WriteLine("Explorer reuse test failed: first open did not create exactly one target tab.");
                DumpShellWindows(firstTargetWindow);
                DumpDebugLog(debugLog);
                return 1;
            }

            await Task.Delay(repeatDelayMs);
            StartExplorer(target);
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

            Console.WriteLine($"PASS Explorer reuse: reopening '{target}' reused the existing tab without creating a duplicate.");
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

            if (defaultWindows.Length == 0 || transparentDefaultWindows.Length > 0)
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
        var observeMs = int.TryParse(GetOption(args, "--observe-ms"), out var parsedObserveMs) ? parsedObserveMs : 1_800;
        var maxHideMs = int.TryParse(GetOption(args, "--max-hide-ms"), out var parsedMaxHideMs) ? parsedMaxHideMs : 900;
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
        var openedByTest = new HashSet<nint>();

        try
        {
            await Task.Delay(startupDelayMs);

            var monitor = HiddenDefaultWindowMonitor.Start(baselineDefaultHwnds);
            StartExplorerDefault();
            await Task.Delay(observeMs);
            await monitor.StopAsync();

            foreach (var hwnd in GetShellWindows()
                         .Where(IsDefaultLocation)
                         .Select(window => (nint)window.Hwnd)
                         .Where(hwnd => !baselineDefaultHwnds.Contains(hwnd)))
            {
                openedByTest.Add(hwnd);
            }

            var sightings = monitor.Sightings.ToArray();
            var offending = sightings.Where(sighting => sighting.HiddenMs > maxHideMs).ToArray();
            if (sightings.Length == 0)
            {
                Console.Error.WriteLine("Explorer user-default test failed: no new default-location window was observed.");
                DumpShellWindows(GetShellWindows());
                DumpDebugLog(debugLog);
                return 1;
            }

            if (offending.Length > 0)
            {
                Console.Error.WriteLine("Explorer user-default test failed.");
                Console.Error.WriteLine($"Threshold={maxHideMs}ms");
                foreach (var sighting in sightings)
                {
                    Console.Error.WriteLine($"HWND={sighting.Hwnd} HiddenMs={sighting.HiddenMs:F1} FirstSeen={sighting.FirstSeen:HH:mm:ss.fff} HiddenAt={(sighting.FirstHiddenAt?.ToString("HH:mm:ss.fff") ?? "none")} ReleasedAt={(sighting.LastHiddenAt?.ToString("HH:mm:ss.fff") ?? "none")}");
                }
                DumpShellWindows(GetShellWindows());
                DumpDebugLog(debugLog);
                return 1;
            }

            Console.WriteLine($"PASS Explorer user-default: new default-location windows stayed visible within {maxHideMs}ms (max observed={sightings.Max(s => s.HiddenMs):F0}ms).");
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
                Directory.Delete(path, recursive: true);
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
                        new HiddenDefaultWindowSighting(hWnd, now, transparent ? now : null, transparent ? now : null),
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
        DateTime? LastHiddenAt)
    {
        public double HiddenMs => FirstHiddenAt is { } start && LastHiddenAt is { } end
            ? (end - start).TotalMilliseconds
            : 0;

        public HiddenDefaultWindowSighting Observe(bool transparent, DateTime observedAt)
        {
            if (!transparent)
                return this;

            return this with
            {
                FirstHiddenAt = FirstHiddenAt ?? observedAt,
                LastHiddenAt = observedAt
            };
        }
    }
}
