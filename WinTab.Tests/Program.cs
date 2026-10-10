using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

Console.InputEncoding = new UTF8Encoding(false);
Console.OutputEncoding = new UTF8Encoding(false);
if (Environment.GetEnvironmentVariable("WINTAB_TEST_TRACE") == "1")
    System.Diagnostics.Trace.Listeners.Add(new System.Diagnostics.TextWriterTraceListener(Console.Error));
// WinTab always logs, by default into the user's own log folder. The tests' hooks write to a file of their own
// instead; the stress tests hand the WinTab they start its own log explicitly.
Environment.SetEnvironmentVariable(WinTab.Hooks.ExplorerDebugLog.FileOverrideVariable,
    System.IO.Path.Combine(System.IO.Path.GetTempPath(), "WinTab.Tests", "tests.log"));

if (args.Length == 3 && args[0] == "--conceal-test-window")
    return WindowSafetyTests.ConcealRecoveryWindow(args[1], args[2]);

if (args.Length == 3 && args[0] == "--desktop-pipe-peer")
    return await DesktopUiProcessTests.RunPeerAsync(args[1], int.Parse(args[2]));

string? filter = null;
string? exclude = null;
for (var i = 0; i < args.Length; i++)
{
    if (StringComparer.OrdinalIgnoreCase.Equals(args[i], "--filter") && i + 1 < args.Length)
        filter = args[++i];
    else if (StringComparer.OrdinalIgnoreCase.Equals(args[i], "--exclude") && i + 1 < args.Length)
        exclude = args[++i];
    else
    {
        Console.Error.WriteLine("Unsupported test argument: " + args[i]);
        return 2;
    }
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
            .Concat(ExplorerSessionCommandTests.All())
            .Concat(ExplorerSessionJournalTests.All())
            .Concat(ExplorerShortcutTests.All())
            .Concat(ShortcutTextBoxTests.All())
            .Concat(DesktopBridgeTests.All())
            .Concat(DesktopUiProcessTests.All())
            .Concat(MainWindowSizingTests.All())
            .Concat(RegistryManagerTests.All())
            .Concat(RecycleBinOpenRegistrationTests.All())
            .Concat(ThemeManagerTests.All())
            .Concat(HookManagerTests.All())
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
            .Concat(NotepadTabAutomationTests.All())
            .Concat(ExplorerTabWheelSwitchTests.All())
            .Concat(ExplorerTabRegistrationTests.All())
            .Concat(ExplorerTabLifetimeTests.All())
            .Concat(ExplorerTabReuseTests.All())
            .Concat(ExplorerReuseSelectionTests.All())
            .Concat(ExplorerNativeFocusActivationTests.All())
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
        var unavailable = 0;
        foreach (var (name, body) in tests)
        {
            try
            {
                await body();
                Console.WriteLine($"PASS {name}");
            }
            catch (TestSkippedException ex)
            {
                unavailable++;
                Console.WriteLine($"SKIP {name}: {ex.Message}");
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

        Console.WriteLine($"{tests.Count - failed - unavailable}/{tests.Count - unavailable} passed" +
            (skipped + unavailable > 0 ? $" ({skipped + unavailable} skipped)" : string.Empty));
        Console.Out.Flush();
        return failed == 0 ? 0 : 1;
    }

    /// <summary>Splits a <c>|</c>-separated substring list, or null when nothing was given.</summary>
    private static string[]? Split(string? value)
    {
        var parts = value?.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return parts is { Length: > 0 } ? parts : null;
    }
}
