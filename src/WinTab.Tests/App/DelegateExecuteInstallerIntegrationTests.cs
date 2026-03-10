using System.IO;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public sealed class DelegateExecuteInstallerIntegrationTests
{
    private const string DelegateExecuteClsid = "FD5BF2CD-0B24-4A80-9AF3-E40F9AFC0001";

    [Fact]
    public void InstallerScript_ShouldRegisterAndUnregisterDelegateExecuteComServer()
    {
        string scriptPath = GetRepoFilePath("installers", "WinTab.iss");
        string script = File.ReadAllText(scriptPath);

        script.Should().Contain("#define DelegateExecuteClsid",
            "installer must declare the DelegateExecute CLSID constant");
        script.Should().Contain(DelegateExecuteClsid,
            "installer must declare the expected DelegateExecute GUID");
        script.Should().Contain(@"Software\Classes\CLSID\{#DelegateExecuteClsid}",
            "installer must register DelegateExecute CLSID under HKCU");
        script.Should().Contain(@"Software\Classes\CLSID\{#DelegateExecuteClsid}\InProcServer32",
            "installer must register COM inproc server");
        script.Should().Contain("WinTab.ShellBridge.comhost.dll",
            "delegate execute host must be present in installer registration");
        script.Should().Contain("ThreadingModel",
            "COM inproc registration must set apartment threading model");
        script.Should().Contain("Apartment",
            "delegate execute COM host is apartment-threaded");
        script.Should().Contain("uninsdeletekey",
            "installer uninstall must remove COM registration keys");
    }

    [Fact]
    public void AppProject_ShouldPublishShellBridgeRuntimeConfigAndDeps()
    {
        string projectPath = GetRepoFilePath("src", "WinTab.App", "WinTab.App.csproj");
        string project = File.ReadAllText(projectPath);

        project.Should().Contain("WinTab.ShellBridge.runtimeconfig.json",
            "COM host activation requires component runtimeconfig at runtime");
        project.Should().Contain("WinTab.ShellBridge.deps.json",
            "COM host activation requires component deps metadata at runtime");
        project.Should().Contain("CopyToPublishDirectory",
            "runtime metadata files must be copied during publish");
    }

    [Fact]
    public void InstallerScript_ShouldRestartExplorerWhenUpdatingShellBridge()
    {
        string scriptPath = GetRepoFilePath("installers", "WinTab.iss");
        string script = File.ReadAllText(scriptPath);

        script.Should().Contain("taskkill.exe",
            "updating the Explorer-loaded ShellBridge requires stopping the owning process");
        script.Should().Contain("/IM explorer.exe /F",
            "installer must forcibly stop Explorer so the ShellBridge DLL is not left stale in Program Files");
        script.Should().Contain("explorer.exe",
            "installer must restart Explorer after the shell bridge files are updated");
        script.Should().Contain("DeinitializeSetup",
            "Explorer must still be restarted if setup exits early after stopping the shell");
    }

    [Fact]
    public void InstallerScript_ShouldRunAppCleanupBeforeUpdatingExistingInstallation()
    {
        string scriptPath = GetRepoFilePath("installers", "WinTab.iss");
        string script = File.ReadAllText(scriptPath);

        script.Should().Contain("StopExistingShellBridgeHostsForUpgrade",
            "the installer should centralize shell bridge upgrade cleanup in one place");
        script.Should().Contain("PrepareToInstall",
            "existing-installation cleanup must happen before file copy starts");
        script.Should().Contain("--wintab-cleanup",
            "the installer should disable interception before replacing shell bridge files in an existing installation");
        script.Should().Contain("ExistingInstallDetected",
            "the cleanup and Explorer restart flow should only run when replacing an existing installation");
    }

    private static string GetRepoFilePath(params string[] parts)
    {
        string current = AppContext.BaseDirectory;

        for (int i = 0; i < 10; i++)
        {
            string candidate = Path.Combine(current, Path.Combine(parts));
            if (File.Exists(candidate))
                return candidate;

            string? parent = Directory.GetParent(current)?.FullName;
            if (string.IsNullOrWhiteSpace(parent))
                break;

            current = parent;
        }

        return Path.Combine(AppContext.BaseDirectory, Path.Combine(parts));
    }
}
