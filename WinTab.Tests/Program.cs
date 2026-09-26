using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

Console.InputEncoding = new UTF8Encoding(false);
Console.OutputEncoding = new UTF8Encoding(false);
// WinTab always logs, by default into the user's own log folder. The tests' hooks write to a file of their own
// instead; the stress tests hand the WinTab they start its own log explicitly.
Environment.SetEnvironmentVariable(WinTab.Hooks.ExplorerDebugLog.FileOverrideVariable,
    System.IO.Path.Combine(System.IO.Path.GetTempPath(), "WinTab.Tests", "tests.log"));

if (args.Length == 3 && args[0] == "--conceal-test-window")
    return WindowSafetyTests.ConcealRecoveryWindow(args[1], args[2]);

if (args.Length > 0 && StringComparer.OrdinalIgnoreCase.Equals(args[0], "--session-restore-stress"))
    return await ExplorerStressTest.RunSessionRestoreAsync(args);
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

string? filter = null;
string? exclude = null;
for (var i = 0; i < args.Length - 1; i++)
{
    if (StringComparer.OrdinalIgnoreCase.Equals(args[i], "--filter"))
        filter = args[i + 1];
    else if (StringComparer.OrdinalIgnoreCase.Equals(args[i], "--exclude"))
        exclude = args[i + 1];
}

return await UnitTestRunner.RunAll(filter, exclude);

internal static class UnitTestRunner
{
    public static async Task<int> RunAll(string? filter = null, string? exclude = null)
    {
        var filters = Split(filter);
        var excludes = Split(exclude);
        var tests = Enumerable.Empty<(string Name, Func<Task> Body)>()
            .Concat(ExplorerLaunchLocationResolverTests.All())
            .Concat(LocationTests.All())
            .Concat(PollingTests.All())
            .Concat(StaTaskSchedulerTests.All())
            .Concat(SettingsStoreTests.All())
            .Concat(ExplorerSessionStoreTests.All())
            .Concat(ExplorerSessionTests.All())
            .Concat(ExplorerSessionLocationTests.All())
            .Concat(ExplorerSessionNativeTests.All())
            .Concat(RegistryManagerTests.All())
            .Concat(ThemeManagerTests.All())
            .Concat(BackgroundWorkTests.All())
            .Concat(BufferedDiagnosticLogTests.All())
            .Concat(WindowSafetyTests.All())
            .Concat(SelectionSnapshotTests.All())
            .Concat(TabSelectionEngineTests.All())
            .Concat(NavigationMiddleClickTests.All())
            .Concat(NavigationInputObserverTests.All())
            .Concat(NavigationNativeSelectionTests.All())
            .Concat(MergeSourceConcealPulseTests.All())
            .Concat(ExplorerTabDoubleClickCloseTests.All())
            .Concat(ExplorerTabWheelSwitchTests.All())
            .Concat(ExplorerTabRegistrationTests.All())
            .Concat(ExplorerTabLifetimeTests.All())
            .Concat(ExplorerTabReuseTests.All())
            .Concat(ExplorerTabActivationTests.All())
            .Concat(ExplorerReuseSelectionTests.All())
            .Concat(ExplorerNativeFocusTests.All())
            .Concat(ExplorerDesktopFlowTests.All())
            .Concat(ExplorerDesktopOpenTests.All())
            .Concat(ExplorerPreloadedFrameTests.All())
            .Concat(ExplorerTabTearOffTests.All())
            .Concat(ExplorerDirectTabTests.All())
            .Concat(TaskbarButtonTests.All())
            .Concat(UpdateReleaseParserTests.All())
            .Concat(UpdateManagerTests.All())
            .Concat(DualKeyDictionaryTests.All())
            .Where(test => filters == null || filters.Any(part => test.Name.Contains(part, StringComparison.OrdinalIgnoreCase)))
            .ToList();

        // Host interference (an installed WinTab or Explorer running on this machine) can make a few
        // Explorer-facing tests fail on a developer's desktop. Excluded tests are reported, never silent.
        var skipped = excludes == null ? 0 : tests.RemoveAll(test =>
            excludes.Any(part => test.Name.Contains(part, StringComparison.OrdinalIgnoreCase)));
        if (skipped > 0)
            Console.WriteLine($"SKIP {skipped} test(s) excluded by: {exclude}");

        if (tests.Count == 0)
        {
            Console.Error.WriteLine($"No tests match: {filter}");
            return 1;
        }

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
                // Reflection wraps the real failure; its message is what tells the reader what went wrong.
                while (ex is System.Reflection.TargetInvocationException { InnerException: { } inner })
                    ex = inner;
                Console.Error.WriteLine($"FAIL {name}: {ex.Message}");
                if (Environment.GetEnvironmentVariable("WINTAB_TEST_TRACE") == "1")
                    Console.Error.WriteLine(ex);
            }
        }

        Console.WriteLine($"{tests.Count - failed}/{tests.Count} passed");
        return failed == 0 ? 0 : 1;
    }

    /// <summary>Splits a <c>|</c>-separated substring list, or null when nothing was given.</summary>
    private static string[]? Split(string? value)
    {
        var parts = value?.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return parts is { Length: > 0 } ? parts : null;
    }
}
