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
    }

    [Fact]
    public void UninstallCleanupHandlerSource_ShouldDeleteUserOverridesToRevealNativeExplorerDefaults()
    {
        string sourcePath = TestRepoPaths.GetFile(["src", "WinTab.App", "Services", "UninstallCleanupHandler.cs"]);
        string source = File.ReadAllText(sourcePath);

        source.Should().Contain("shell?.DeleteValue(string.Empty, throwOnMissingValue: false);",
            "fallback cleanup should remove the HKCU default verb override");
        source.Should().Contain("classesRoot.DeleteSubKeyTree($@\"{cls}\\shell\\{verb}\", throwOnMissingSubKey: false);",
            "fallback cleanup should delete the HKCU open/explore/opennewwindow overrides");
        source.Should().Contain("TryDeleteEmptyKey(classesRoot, $@\"{cls}\\shell\");",
            "fallback cleanup should clean up empty HKCU shell keys after removing WinTab overrides");
    }
}
