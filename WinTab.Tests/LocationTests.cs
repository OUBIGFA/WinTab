using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Interop;

internal static class LocationTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("NormalizeLocation converts file URLs to local filesystem paths", NormalizesFileUrlsToLocalPaths);
        yield return ("NormalizeLocation keeps web URLs usable", KeepsWebUrlsUsable);
        yield return ("NormalizeLocation keeps the UNC prefix while trimming the trailing separator", KeepsUncPrefix);
        yield return ("NormalizeLocation trims quotes and trailing separators from local paths", TrimsQuotesAndSeparators);
        yield return ("NormalizeLocation expands environment variables", ExpandsEnvironmentVariables);
        yield return ("NormalizeLocation preserves absolute drive roots and rooted paths", PreservesRoots);
        yield return ("NormalizeLocation prefixes bare shell CLSID forms with shell:", PrefixesShellClsidForms);
        yield return ("Recycle Bin and virtual folder PIDL resolution and equivalence", RecycleBinAndVirtualFolderPidlEquivalence);
        yield return ("Different real folders are not reported as equivalent", DifferentFoldersAreNotEquivalent);
    }

    private static Task NormalizesFileUrlsToLocalPaths()
    {
        Check.EqualIgnoreCase(@"C:\Users\Public\Downloads", Helper.NormalizeLocation("file:///C:/Users/Public/Downloads"));
        Check.EqualIgnoreCase(@"C:\Users\Public\Downloads", Helper.NormalizeLocation("file:///C:/Users/Public/Downloads/"));
        return Task.CompletedTask;
    }

    private static Task KeepsWebUrlsUsable()
    {
        Check.Equal("https://example.com/path/to/file", Helper.NormalizeLocation("https://example.com/path/to/file"));
        Check.Equal("www.example.com/path", Helper.NormalizeLocation("www.example.com/path"));
        return Task.CompletedTask;
    }

    private static Task KeepsUncPrefix()
    {
        Check.Equal(@"\\server\share\folder", Helper.NormalizeLocation(@"\\server\share\folder\"));
        Check.Equal(@"\\server\share", Helper.NormalizeLocation("//server/share/"));
        return Task.CompletedTask;
    }

    private static Task TrimsQuotesAndSeparators()
    {
        Check.Equal(@"D:\Data\Photos", Helper.NormalizeLocation("\"D:\\Data\\Photos\\\""));
        Check.Equal(@"D:\Data\Photos", Helper.NormalizeLocation(" D:/Data/Photos/ "));
        return Task.CompletedTask;
    }

    private static Task ExpandsEnvironmentVariables()
    {
        var expectedRoot = Environment.GetEnvironmentVariable("SystemRoot")!;
        Check.EqualIgnoreCase(expectedRoot + @"\System32", Helper.NormalizeLocation(@"%SystemRoot%\System32\"));
        return Task.CompletedTask;
    }

    private static Task PreservesRoots()
    {
        Check.Equal(@"C:\", Helper.NormalizeLocation(@"C:\"));
        Check.Equal(@"D:\", Helper.NormalizeLocation("D:/"));
        Check.Equal(@"C:\", Helper.NormalizeLocation("file:///C:/"));
        Check.Equal(@"\folder", Helper.NormalizeLocation(@"\folder\"));
        Check.Equal(@"C:folder", Helper.NormalizeLocation(@"C:folder\"));
        return Task.CompletedTask;
    }

    private static Task PrefixesShellClsidForms()
    {
        const string guid = "{645FF040-5081-101B-9F08-00AA002F954E}";
        Check.Equal($"shell:::{guid}", Helper.NormalizeLocation($"::{guid}"));
        Check.Equal($"shell:::{guid}", Helper.NormalizeLocation(guid));
        Check.Equal($"shell:::{guid}", Helper.NormalizeLocation($"shell:::{guid}"));
        return Task.CompletedTask;
    }

    private static Task RecycleBinAndVirtualFolderPidlEquivalence()
    {
        using var comparer = new ShellPathComparer();

        const string recycleBinPath = "shell:::{645FF040-5081-101B-9F08-00AA002F954E}";
        const string cleanRecycleBinPath = "::{645FF040-5081-101B-9F08-00AA002F954E}";
        const string shortcutRecycleBinPath = "shell:RecycleBinFolder";
        var normalizedUnc = Helper.NormalizeLocation(@"\\127.0.0.1\c$");

        var pidl1 = comparer.GetPidlFromPath(recycleBinPath);
        var pidl2 = comparer.GetPidlFromPath(cleanRecycleBinPath);
        var pidl3 = comparer.GetPidlFromPath(shortcutRecycleBinPath);
        var pidlUnc = comparer.GetPidlFromPath(normalizedUnc);

        try
        {
            Check.That(pidl1 != 0, "PIDL for shell:::RecycleBin path should be resolved.");
            Check.That(pidl2 != 0, "PIDL for clean RecycleBin path should be resolved.");
            Check.That(pidl3 != 0, "PIDL for shell:RecycleBinFolder path should be resolved.");
            Check.That(pidlUnc != 0, $"PIDL for normalized UNC path '{normalizedUnc}' should be resolved.");

            Check.That(comparer.IsEquivalent(pidl1, pidl2), "Virtual path and clean virtual path should be equivalent.");
            Check.That(comparer.IsEquivalent(recycleBinPath, cleanRecycleBinPath), "Virtual path string comparisons should be equivalent.");
            Check.That(comparer.IsEquivalent(shortcutRecycleBinPath, cleanRecycleBinPath), "Shortcut and clean path should be equivalent.");
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

    private static Task DifferentFoldersAreNotEquivalent()
    {
        using var comparer = new ShellPathComparer();
        var windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        var system32 = Environment.GetFolderPath(Environment.SpecialFolder.System);

        Check.That(comparer.IsEquivalent(windows, windows + @"\"), "The same folder with and without a trailing separator must be equivalent.");
        Check.That(!comparer.IsEquivalent(windows, system32), "Distinct folders must not be reported as equivalent.");
        Check.That(!comparer.IsEquivalent(@"Z:\definitely\missing\one", @"Z:\definitely\missing\two"), "Two unparseable paths must not be reported as equivalent.");
        return Task.CompletedTask;
    }
}
