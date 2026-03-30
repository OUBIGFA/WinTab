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
}
