using System.IO;
using System.Runtime.CompilerServices;

namespace WinTab.Tests.App;

internal static class TestRepoPaths
{
    private const string RepoRootEnvVar = "WINTAB_REPO_ROOT";

    public static string GetFile(string[] parts, [CallerFilePath] string callerFilePath = "")
    {
        string repoRoot = FindRepoRoot(callerFilePath)
            ?? throw new DirectoryNotFoundException("Unable to locate the WinTab repository root for source-based tests.");

        return Path.Combine(repoRoot, Path.Combine(parts));
    }

    private static string? FindRepoRoot(string callerFilePath)
    {
        string? envRepoRoot = Environment.GetEnvironmentVariable(RepoRootEnvVar);
        if (!string.IsNullOrWhiteSpace(envRepoRoot) && IsRepoRoot(envRepoRoot))
            return envRepoRoot;

        string[] seeds =
        [
            Path.GetDirectoryName(callerFilePath) ?? string.Empty,
            AppContext.BaseDirectory,
            Directory.GetCurrentDirectory()
        ];

        foreach (string seed in seeds)
        {
            string? current = seed;
            for (int i = 0; i < 20 && !string.IsNullOrWhiteSpace(current); i++)
            {
                if (IsRepoRoot(current))
                    return current;

                current = Directory.GetParent(current)?.FullName;
            }
        }

        return null;
    }

    private static bool IsRepoRoot(string path)
    {
        return File.Exists(Path.Combine(path, "WinTab.slnx")) &&
               Directory.Exists(Path.Combine(path, "src")) &&
               Directory.Exists(Path.Combine(path, "installers"));
    }
}
