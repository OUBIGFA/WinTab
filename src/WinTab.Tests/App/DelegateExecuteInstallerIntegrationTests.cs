using System.IO;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public sealed class DelegateExecuteInstallerIntegrationTests
{
    private const string DelegateExecuteClsid = "FD5BF2CD-0B24-4A80-9AF3-E40F9AFC0001";

    [Fact]
    public void InstallerScript_ShouldDeclareAndCleanupLegacyDelegateExecuteComServer()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        script.Should().Contain("#define DelegateExecuteClsid",
            "installer must declare the DelegateExecute CLSID constant");
        script.Should().Contain(DelegateExecuteClsid,
            "installer must declare the expected DelegateExecute GUID");
        script.Should().Contain(@"Software\Classes\CLSID\{#DelegateExecuteClsid}",
            "installer must still know the CLSID so uninstall cleanup can remove legacy registrations");
        script.Should().Contain("WinTabDelegateExecuteCleanupHKLM64",
            "installer uninstall must remove old machine-wide DelegateExecute registration keys");
        script.Should().Contain("WinTabDelegateExecuteCleanup32",
            "installer uninstall must remove old per-user 32-bit DelegateExecute registration keys too");
    }

    [Fact]
    public void InstallerScript_ShouldNotPreRegisterPersistentDelegateExecuteComServer()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        script.Should().NotContain(@"Root: HKLM; Subkey: ""Software\Classes\CLSID\{#DelegateExecuteClsid}""",
            "the safety-first redesign must not leave machine-wide DelegateExecute registration behind when WinTab is installed but not running");
        script.Should().NotContain(@"Root: HKLM32; Subkey: ""Software\Classes\CLSID\{#DelegateExecuteClsid}""",
            "the installer must not pre-register the 32-bit machine-wide DelegateExecute bridge either");
        script.Should().NotContain(@"Root: HKCU64; Subkey: ""Software\Classes\CLSID\{#DelegateExecuteClsid}""",
            "runtime-only interception must not rely on persistent per-user COM registration written at install time");
        script.Should().NotContain(@"Root: HKCU32; Subkey: ""Software\Classes\CLSID\{#DelegateExecuteClsid}""",
            "runtime-only interception must not rely on persistent 32-bit per-user COM registration written at install time");
    }

    [Fact]
    public void InstallerScript_ShouldCaptureExplorerShellBaselineDuringInstall()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        script.Should().Contain("--wintab-capture-shell-baseline",
            "install should capture the native Explorer shell baseline immediately so later cleanup can restore the machine's original values instead of hard-coded defaults");
    }

    [Fact]
    public void InstallerScript_ShouldNotRegisterDelegateExecuteForBothRegistryViewsAtInstallTime()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        script.Should().NotContain("Root: HKLM; Subkey: \"Software\\Classes\\CLSID\\{#DelegateExecuteClsid}\"",
            "the installer must not pre-register machine-wide DelegateExecute keys while WinTab is not running");
        script.Should().NotContain("Root: HKCU64; Subkey: \"Software\\Classes\\CLSID\\{#DelegateExecuteClsid}\"",
            "the installer must not pre-register per-user DelegateExecute keys while WinTab is not running");
        script.Should().Contain("/reg:32",
            "uninstall cleanup should explicitly remove the 32-bit DelegateExecute registration too");
    }

    [Fact]
    public void InstallerScript_ShouldCleanupMachineWideDelegateExecuteRegistrationOnUninstall()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        script.Should().Contain("WinTabDelegateExecuteCleanupHKLM64",
            "uninstall must include a dedicated cleanup step for the machine-wide 64-bit COM registration");
        script.Should().Contain("WinTabDelegateExecuteCleanupHKLM32",
            "uninstall must include a dedicated cleanup step for the machine-wide 32-bit COM registration");
        script.Should().Contain(@"HKLM\Software\Classes\CLSID\{#DelegateExecuteClsid}",
            "uninstall must target the machine-wide DelegateExecute CLSID path");
        script.Should().Contain(@"/reg:32",
            "uninstall must explicitly clean up the 32-bit machine-wide registration too");
    }

    [Fact]
    public void AppProject_ShouldPublishShellBridgeRuntimeConfigAndDeps()
    {
        string projectPath = TestRepoPaths.GetFile(["src", "WinTab.App", "WinTab.App.csproj"]);
        string project = File.ReadAllText(projectPath);

        project.Should().Contain("WinTab.ShellBridge.runtimeconfig.json",
            "COM host activation requires component runtimeconfig at runtime");
        project.Should().Contain("WinTab.ShellBridge.deps.json",
            "COM host activation requires component deps metadata at runtime");
        project.Should().Contain("CopyToPublishDirectory",
            "runtime metadata files must be copied during publish");
    }

    [Fact]
    public void AppProject_ShouldPublishShellBridgeX86SidecarFor32BitShellHosts()
    {
        string projectPath = TestRepoPaths.GetFile(["src", "WinTab.App", "WinTab.App.csproj"]);
        string project = File.ReadAllText(projectPath);

        project.Should().Contain("win-x86",
            "the app publish pipeline must produce an x86 ShellBridge sidecar so 32-bit shell hosts can activate DelegateExecute");
        project.Should().Contain("TargetFramework=net8.0-windows",
            "the x86 ShellBridge sidecar must target a runtime family that is broadly available on 32-bit hosts instead of hard-requiring x86 .NET 9");
        project.Should().Contain("$(PublishDir)x86\\",
            "the x86 ShellBridge publish output should live under a dedicated x86 subfolder next to the main app payload");
        project.Should().Contain("PublishShellBridgeX86ForDelegateExecute",
            "publishing the app should automatically include the x86 ShellBridge sidecar required for shell-host compatibility");
    }

    [Fact]
    public void ShellBridgeProject_ShouldDeclareBothRuntimeIdentifiersForDelegateExecute()
    {
        string projectPath = TestRepoPaths.GetFile(["src", "WinTab.ShellBridge", "WinTab.ShellBridge.csproj"]);
        string project = File.ReadAllText(projectPath);

        project.Should().Contain("<RuntimeIdentifiers>win-x64;win-x86</RuntimeIdentifiers>",
            "the ShellBridge project must restore assets for both architectures before the app publish target requests the x86 sidecar");
        project.Should().Contain("<TargetFrameworks>net8.0-windows;net9.0-windows</TargetFrameworks>",
            "the ShellBridge must multi-target so the x86 sidecar can avoid depending on x86 .NET 9 while the main app remains on net9");
        project.Should().Contain("<RollForward>LatestMajor</RollForward>",
            "the DelegateExecute bridge should accept newer installed runtimes instead of failing when the exact x86 runtime patch line is absent");
    }

    [Fact]
    public void ShellBridgeDependencies_ShouldMultiTargetForNet8CompatibleX86Sidecar()
    {
        string coreProjectPath = TestRepoPaths.GetFile(["src", "WinTab.Core", "WinTab.Core.csproj"]);
        string diagnosticsProjectPath = TestRepoPaths.GetFile(["src", "WinTab.Diagnostics", "WinTab.Diagnostics.csproj"]);
        string win32ProjectPath = TestRepoPaths.GetFile(["src", "WinTab.Platform.Win32", "WinTab.Platform.Win32.csproj"]);

        File.ReadAllText(coreProjectPath).Should().Contain("<TargetFrameworks>net8.0;net9.0</TargetFrameworks>",
            "the x86 ShellBridge cannot target net8 unless its shared core dependency also multi-targets");
        File.ReadAllText(diagnosticsProjectPath).Should().Contain("<TargetFrameworks>net8.0;net9.0</TargetFrameworks>",
            "the Win32 shell interop layer depends on diagnostics and must stay buildable for the net8 x86 sidecar");
        File.ReadAllText(win32ProjectPath).Should().Contain("<TargetFrameworks>net8.0-windows;net9.0-windows</TargetFrameworks>",
            "the platform interop layer must multi-target so the x86 ShellBridge can build against net8 while the app stays on net9");
    }

    [Fact]
    public void InstallerScript_ShouldNotForceRestartExplorerWhenUpdatingShellBridge()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        script.Should().NotContain("/IM explorer.exe /F",
            "setup should not kill the user's shell during install or uninstall");
    }

    [Fact]
    public void InstallerScript_ShouldRunAppCleanupBeforeUpdatingExistingInstallation()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
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
}
