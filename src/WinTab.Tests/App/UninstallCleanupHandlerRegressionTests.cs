using System.IO;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public sealed class UninstallCleanupHandlerRegressionTests
{
    [Fact]
    public void UninstallCleanupHandlerSource_ShouldAttemptToRemoveMachineWideAndUserScopeDelegateExecuteRegistration()
    {
        string sourcePath = TestRepoPaths.GetFile(["src", "WinTab.App", "Services", "UninstallCleanupHandler.cs"]);
        string source = File.ReadAllText(sourcePath);

        source.Should().Contain("RegistryHive.LocalMachine",
            "uninstall cleanup must attempt to remove machine-wide COM registration introduced for Windows 11 Start Menu and elevated shell hosts");
        source.Should().Contain("RegistryHive.CurrentUser",
            "uninstall cleanup must continue removing user-scope registration too");
        source.Should().Contain("Failed to clean up HKLM COM registration",
            "machine-wide cleanup failures should be logged explicitly for diagnosis");
        source.Should().Contain("DelegateExecuteClsidBraced",
            "cleanup must target the WinTab DelegateExecute CLSID consistently");
        source.Should().Contain("MalformedDelegateExecuteClsidBraced",
            "cleanup should also target the malformed legacy CLSID key written by the broken installer so upgraded machines recover automatically");
    }

    [Fact]
    public void UninstallCleanupHandlerSource_ShouldDeleteUserOverridesToRevealNativeExplorerDefaults()
    {
        string sourcePath = TestRepoPaths.GetFile(["src", "WinTab.App", "Services", "UninstallCleanupHandler.cs"]);
        string source = File.ReadAllText(sourcePath);

        source.Should().Contain("shell?.DeleteValue(string.Empty, throwOnMissingValue: false);",
            "standalone cleanup should remove the HKCU default verb override so Explorer can fall back to the native Win11 behavior");
        source.Should().Contain("classesRoot.DeleteSubKeyTree($@\"{cls}\\shell\\{verb}\", throwOnMissingSubKey: false);",
            "fallback cleanup should delete the HKCU open/explore/opennewwindow overrides");
        source.Should().Contain("TryDeleteEmptyKey(classesRoot, $@\"{cls}\\shell\");",
            "fallback cleanup should clean up empty HKCU shell keys after removing WinTab overrides");
        source.Should().Contain("SHChangeNotify",
            "after removing shell overrides, cleanup must notify the shell so Explorer drops any stale association cache without requiring the user to log off");
    }
}
