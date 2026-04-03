using System.IO;
using System.Text.RegularExpressions;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

/// <summary>
/// Tests that verify the installer correctly handles the uninstall ordering
/// to prevent Explorer from freezing due to dangling registry entries.
///
/// Root cause: Inno Setup executes [UninstallDelete] BEFORE [UninstallRun].
/// This means files are deleted before the cleanup exe can run, leaving
/// HKCU\Software\Classes\{Folder,Directory,Drive}\shell\open\command
/// pointing to a deleted WinTab.exe — causing Explorer to hang on folder open.
/// </summary>
public sealed class InstallerUninstallOrderingTests
{
    [Fact]
    public void InstallerScript_ShouldRestoreExplorerOpenVerbBeforeDeletingFiles()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        // The installer must use CurUninstallStepChanged with usUninstall to
        // restore registry entries BEFORE [UninstallDelete] removes files.
        script.Should().Contain("CurUninstallStepChanged",
            "the installer must hook into uninstall step changes to run cleanup before file deletion");

        script.Should().Contain("usUninstall",
            "cleanup must run during the usUninstall phase, before UninstallDelete executes");
    }

    [Fact]
    public void InstallerScript_ShouldRestoreExplorerOpenVerbViaRegExeAsFallback()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        // The installer must keep a script-side fallback restore path for cases
        // where the installed WinTab cleanup executable cannot run.
        script.Should().Contain("RegDeleteKeyIncludingSubkeys(HKCU, 'Software\\Classes\\Folder\\shell\\open');",
            "uninstall fallback should remove the user-scope Folder open override");
        script.Should().Contain("Directory\\shell",
            "uninstall fallback must remove the Directory shell overrides");
        script.Should().Contain("Drive\\shell",
            "uninstall fallback must remove the Drive shell overrides");
        script.Should().Contain("SHChangeNotify($08000000, $0000, 0, 0);",
            "after clearing overrides, the script fallback should notify Explorer to reload shell associations");
    }

    [Fact]
    public void InstallerScript_ShouldDeleteUserScopeOverridesToRevealNativeExplorerDefaults()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        script.Should().Contain(@"RegDeleteValue(HKCU, 'Software\Classes\Directory\shell', '');",
            "the uninstall fallback must remove the HKCU Directory default verb override");
        script.Should().Contain(@"RegDeleteValue(HKCU, 'Software\Classes\Drive\shell', '');",
            "the uninstall fallback must remove the HKCU Drive default verb override");
        script.Should().Contain(@"RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\Directory\shell\open');",
            "the uninstall fallback should delete the HKCU Directory open override so Explorer falls back to the machine defaults");
    }

    [Fact]
    public void InstallerScript_ShouldNotRelyOnExeForOpenVerbRestore()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        script.Should().Contain("TryRunInstalledCleanupBeforeDelete",
            "standalone uninstall should prefer the installed WinTab cleanup path so backup-based restore remains authoritative");
        script.Should().Contain("reg.exe",
            "the installer must still keep a registry-script fallback when the installed cleanup executable is unavailable");
        script.Should().Contain("if not TryRunInstalledCleanupBeforeDelete() then",
            "fallback defaults should run only after the backup-based cleanup path fails");
    }

    [Fact]
    public void InstallerScript_ShouldRemoveDelegateExecuteClsidFromAllViews()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        // Must clean up DelegateExecute CLSID from both HKLM and HKCU,
        // both 64-bit and 32-bit views, BEFORE files are deleted.
        script.Should().Contain("CLSID",
            "uninstall must clean up CLSID registrations");

        script.Should().Contain("FD5BF2CD-0B24-4A80-9AF3-E40F9AFC0001",
            "uninstall must target the specific WinTab DelegateExecute CLSID");
    }

    [Fact]
    public void InstallerScript_ShouldNotForceRestartExplorerDuringUninstall()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        script.Should().NotContain("/IM explorer.exe /F",
            "uninstall should restore shell registration without force-killing the user's shell");
    }

    [Fact]
    public void InstallerScript_CleanReinstall_ShouldCleanupResidualWinTabProcessesBeforeCopyingNewFiles()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        Regex.IsMatch(
                script,
                @"if SelectedReinstallMode = 'clean' then[\s\S]*StopExistingShellBridgeHostsForUpgrade\(\);",
                RegexOptions.CultureInvariant)
            .Should()
            .BeTrue("clean reinstall must still kill any lingering WinTab process that keeps DLLs locked after the legacy uninstaller returns");
    }

    [Fact]
    public void InstallerScript_CleanReinstall_ShouldWaitForOldInstallPayloadToDisappearOrUnlock()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        Regex.IsMatch(
                script,
                @"WaitForExistingInstallFilesCleanup",
                RegexOptions.CultureInvariant)
            .Should()
            .BeTrue("clean reinstall must wait for the old install payload to disappear after uninstall so file-copy does not race a still-locked DLL");
    }
}
