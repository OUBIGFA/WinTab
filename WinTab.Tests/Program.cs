using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

Console.InputEncoding = new UTF8Encoding(false);
Console.OutputEncoding = new UTF8Encoding(false);

if (args.Length == 3 && args[0] == "--conceal-test-window")
    return WindowSafetyTests.ConcealRecoveryWindow(args[1], args[2]);

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

return await UnitTestRunner.RunAll();

internal static class UnitTestRunner
{
    public static async Task<int> RunAll()
    {
        var tests = Enumerable.Empty<(string Name, Func<Task> Body)>()
            .Concat(ExplorerLaunchLocationResolverTests.All())
            .Concat(LocationTests.All())
            .Concat(PollingTests.All())
            .Concat(StaTaskSchedulerTests.All())
            .Concat(SettingsStoreTests.All())
            .Concat(ThemeManagerTests.All())
            .Concat(BackgroundWorkTests.All())
            .Concat(BufferedDiagnosticLogTests.All())
            .Concat(WindowSafetyTests.All())
            .Concat(SelectionSnapshotTests.All())
            .Concat(TabSelectionEngineTests.All())
            .Concat(MergeSourceConcealPulseTests.All())
            .Concat(ExplorerTabDoubleClickCloseTests.All())
            .Concat(ExplorerTabRegistrationTests.All())
            .Concat(ExplorerTabLifetimeTests.All())
            .Concat(ExplorerTabReuseTests.All())
            .Concat(ExplorerTabActivationTests.All())
            .Concat(ExplorerReuseSelectionTests.All())
            .Concat(ExplorerDesktopFlowTests.All())
            .Concat(UpdateReleaseParserTests.All())
            .Concat(UpdateManagerTests.All())
            .Concat(DualKeyDictionaryTests.All())
            .ToList();

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

        Console.WriteLine($"{tests.Count - failed}/{tests.Count} passed");
        return failed == 0 ? 0 : 1;
    }
}
