using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using WinTab.Managers;

internal static class UpdateReleaseParserTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("update asset selection prefers the installer for each architecture", PrefersMatchingArchitectureInstaller);
        yield return ("update asset selection falls back to the first installer for unknown architectures", FallsBackToFirstInstaller);
        yield return ("update asset selection ignores non-installer assets", IgnoresNonInstallerAssets);
        yield return ("release tags normalize to comparable versions", NormalizesReleaseTags);
    }

    private static Task PrefersMatchingArchitectureInstaller()
    {
        var release = BuildRelease(
            ("WinTab_v1.2.0_x86_Setup.exe", "https://dl/x86"),
            ("WinTab_v1.2.0_arm64_Setup.exe", "https://dl/arm64"),
            ("WinTab_v1.2.0_x64_Setup.exe", "https://dl/x64"));

        Check.Equal("https://dl/x64", UpdateReleaseParser.FindMatchingAssetUrl(release, Architecture.X64), "x64 process");
        Check.Equal("https://dl/x86", UpdateReleaseParser.FindMatchingAssetUrl(release, Architecture.X86), "x86 process");
        Check.Equal("https://dl/arm64", UpdateReleaseParser.FindMatchingAssetUrl(release, Architecture.Arm64), "arm64 process");
        return Task.CompletedTask;
    }

    private static Task FallsBackToFirstInstaller()
    {
        var release = BuildRelease(
            ("WinTab_v1.2.0_x86_Setup.exe", "https://dl/x86"),
            ("WinTab_v1.2.0_x64_Setup.exe", "https://dl/x64"));

        Check.Equal("https://dl/x86", UpdateReleaseParser.FindMatchingAssetUrl(release, Architecture.Arm64), "no arm64 asset");
        Check.Equal("https://dl/x86", UpdateReleaseParser.FindMatchingAssetUrl(release, Architecture.Wasm), "unknown architecture");
        return Task.CompletedTask;
    }

    private static Task IgnoresNonInstallerAssets()
    {
        var release = BuildRelease(
            ("WinTab_v1.2.0_x64.zip", "https://dl/zip"),
            ("checksums.txt", "https://dl/sums"));

        Check.Equal<string?>(null, UpdateReleaseParser.FindMatchingAssetUrl(release, Architecture.X64), "no installer");
        Check.Equal<string?>(null, UpdateReleaseParser.FindMatchingAssetUrl(new JsonObject(), Architecture.X64), "no assets array");
        return Task.CompletedTask;
    }

    private static Task NormalizesReleaseTags()
    {
        Check.That(UpdateReleaseParser.TryNormalizeVersion("v1.2.3", out var plain), "plain tag must parse");
        Check.Equal(new Version(1, 2, 3, 0), plain);

        Check.That(UpdateReleaseParser.TryNormalizeVersion("V2.0-beta.1", out var prerelease), "pre-release tag must parse");
        Check.Equal(new Version(2, 0, 0, 0), prerelease);

        Check.That(UpdateReleaseParser.TryNormalizeVersion(" 1.0.1+build7 ", out var build), "build metadata must be stripped");
        Check.Equal(new Version(1, 0, 1, 0), build);

        Check.That(!UpdateReleaseParser.TryNormalizeVersion("latest", out _), "non-numeric tag must be rejected");

        Check.Equal(new Version(1, 0, 0, 0), UpdateReleaseParser.NormalizeVersion(new Version(1, 0)), "missing components must become zero");
        Check.That(UpdateReleaseParser.NormalizeVersion(new Version(1, 0, 1)) > UpdateReleaseParser.NormalizeVersion(new Version(1, 0)),
            "1.0.1 must compare newer than 1.0 once both are normalized");
        return Task.CompletedTask;
    }

    private static JsonObject BuildRelease(params (string Name, string Url)[] assets)
    {
        var array = new JsonArray();
        foreach (var (name, url) in assets)
            array.Add(new JsonObject { ["name"] = name, ["browser_download_url"] = url });

        return new JsonObject { ["tag_name"] = "v1.2.0", ["assets"] = array };
    }
}
