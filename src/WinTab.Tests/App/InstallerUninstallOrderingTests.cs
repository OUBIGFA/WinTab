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

        // The installer must restore open-verb defaults during uninstall.
        // It does this via RegDeleteKeyIncludingSubkeys in RestoreExplorerOpenVerbDefaultsViaReg,
        // which removes the entire command subkeys for all target classes.
        script.Should().Contain("RegDeleteKeyIncludingSubkeys",
            "uninstall must remove Explorer open-verb command keys via RegDeleteKeyIncludingSubkeys");

        script.Should().Contain("Folder\\shell\\open\\command",
            "uninstall must restore Folder shell open verb");

        script.Should().Contain("Directory\\shell",
            "uninstall must restore Directory shell default verb");

        script.Should().Contain("Drive\\shell",
            "uninstall must restore Drive shell default verb");

        script.Should().Contain("RegWriteStringValue",
            "uninstall must write safe default verb values after removing overrides");
    }

    [Fact]
    public void InstallerScript_ShouldNotRelyOnExeForOpenVerbRestore()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        // The open-verb restore must NOT depend on WinTab.exe being present,
        // because [UninstallDelete] removes it before [UninstallRun] executes.
        // Instead, use reg.exe commands in CurUninstallStepChanged.

        // Verify reg.exe is used for open-verb restoration (not just exe cleanup)
        script.Should().Contain("reg.exe",
            "uninstall must use reg.exe to restore Explorer open-verb defaults as a fallback");
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
    public void InstallerScript_CleanReinstall_ShouldNotKillExplorerAgainAfterLegacyUninstallCompletes()
    {
        string scriptPath = TestRepoPaths.GetFile(["installers", "WinTab.iss"]);
        string script = File.ReadAllText(scriptPath);

        Regex.IsMatch(
                script,
                @"if SelectedReinstallMode = 'direct' then\s+StopExistingShellBridgeHostsForUpgrade\(\);",
                RegexOptions.CultureInvariant)
            .Should()
            .BeTrue("clean reinstall already runs the legacy uninstaller, which performs its own Explorer restart");

        Regex.IsMatch(
                script,
                @"if Result <> '' then\s+exit;\s+StopExistingShellBridgeHostsForUpgrade\(\);",
                RegexOptions.CultureInvariant)
            .Should()
            .BeFalse("calling StopExistingShellBridgeHostsForUpgrade unconditionally after clean uninstall kills Explorer twice");
    }
}
