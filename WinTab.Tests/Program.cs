using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;

if (args.Length > 0 && StringComparer.OrdinalIgnoreCase.Equals(args[0], "--stress"))
    return await ExplorerStressTest.RunAsync(args);

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
            ("keeps web URLs usable", KeepsWebUrlsUsable)
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

internal static class ExplorerStressTest
{
    public static async Task<int> RunAsync(string[] args)
    {
        var appPath = GetOption(args, "--app");
        if (string.IsNullOrWhiteSpace(appPath) || !File.Exists(appPath))
            throw new InvalidOperationException("Pass --app with the WinTab.exe path to stress-test.");

        var roundCount = int.TryParse(GetOption(args, "--rounds"), out var parsedRounds) ? parsedRounds : 7;
        var intervalMs = int.TryParse(GetOption(args, "--interval"), out var parsedInterval) ? parsedInterval : 75;
        var startupDelayMs = int.TryParse(GetOption(args, "--startup-delay"), out var parsedStartupDelay) ? parsedStartupDelay : 3_000;
        var root = Path.Combine(Path.GetTempPath(), "WinTabExplorerStress", DateTime.Now.ToString("yyyyMMddHHmmssfff"));
        var debugLog = Path.Combine(root, "wintab-debug.log");
        var targets = CreateTargets(root, roundCount);
        var before = GetShellWindows();
        var beforeDefaultCount = before.Count(IsDefaultLocation);

        using var app = StartWinTab(appPath, debugLog);
        try
        {
            await Task.Delay(startupDelayMs);

            foreach (var target in targets)
            {
                StartExplorer(target);
                await Task.Delay(intervalMs);
            }

            var finalWindows = await WaitForTargetsAsync(targets, beforeDefaultCount);
            var missingTargets = targets.Where(target => finalWindows.All(window => !IsSameFolder(window, target))).ToArray();
            var defaultCount = finalWindows.Count(IsDefaultLocation);

            if (missingTargets.Length > 0 || defaultCount > beforeDefaultCount)
            {
                Console.Error.WriteLine("Explorer stress test failed.");
                Console.Error.WriteLine($"Missing targets: {string.Join(", ", missingTargets)}");
                Console.Error.WriteLine($"Default-location windows before={beforeDefaultCount}, after={defaultCount}");
                DumpShellWindows(finalWindows);
                DumpDebugLog(debugLog);
                return 1;
            }

            Console.WriteLine($"PASS Explorer stress: {targets.Length} rapid folder opens resolved to target folders.");
            Console.WriteLine($"PASS Default-location windows did not increase: before={beforeDefaultCount}, after={defaultCount}.");
            return 0;
        }
        finally
        {
            CloseTestShellWindows(root);
            TryKill(app);
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

    private static Process StartWinTab(string appPath, string debugLog)
    {
        foreach (var existingProcess in Process.GetProcessesByName("WinTab"))
            TryKill(existingProcess);

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

    private static async Task<IReadOnlyList<ShellWindowSnapshot>> WaitForTargetsAsync(string[] targets, int beforeDefaultCount)
    {
        var timeoutAt = Environment.TickCount64 + 14_000;
        var last = GetShellWindows();
        while (Environment.TickCount64 < timeoutAt)
        {
            last = GetShellWindows();
            var allTargetsFound = targets.All(target => last.Any(window => IsSameFolder(window, target)));
            var defaultCount = last.Count(IsDefaultLocation);
            if (allTargetsFound && defaultCount <= beforeDefaultCount)
                return last;

            await Task.Delay(200);
        }

        return last;
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

        return new ShellWindowSnapshot(hwnd, name, url, path);
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

    private static void DumpShellWindows(IReadOnlyList<ShellWindowSnapshot> windows)
    {
        foreach (var window in windows)
            Console.Error.WriteLine($"HWND={window.Hwnd} Name='{window.LocationName}' Url='{window.LocationUrl}' Path='{window.Path}'");
    }

    private static void DumpDebugLog(string debugLog)
    {
        if (!File.Exists(debugLog))
            return;

        Console.Error.WriteLine("WinTab debug log:");
        foreach (var line in File.ReadLines(debugLog).TakeLast(200))
            Console.Error.WriteLine(line);
    }

    private static void ReleaseComObject(object? value)
    {
        if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
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

    private sealed record ShellWindowSnapshot(long Hwnd, string LocationName, string LocationUrl, string Path);
}
