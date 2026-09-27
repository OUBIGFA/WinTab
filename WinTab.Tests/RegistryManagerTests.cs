using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Win32;
using WinTab.Helpers;
using WinTab.Managers;

internal static class RegistryManagerTests
{
    private const string RunPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ApprovedPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run";

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("startup reads do not create missing registry keys", MissingKeysStayMissingOnRead);
        yield return ("startup can be enabled and disabled when registry keys do not exist", MissingKeysAreCreatedOnEnable);
        yield return ("startup re-enables an existing disabled approval entry", DisabledApprovalIsEnabled);
        yield return ("Windows' folder-window restore setting is read without creating keys", FolderWindowRestoreSetting);
    }

    private static Task FolderWindowRestoreSetting() => WithRoot(root =>
    {
        const string advanced = @"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced";
        Check.That(!RegistryManager.RestoresFolderWindowsAtSignInUnder(root), "The option is off in a new profile.");
        Check.Equal(0, root.SubKeyCount, "Reading the option must not create Explorer's keys.");
        using var key = root.CreateSubKey(advanced);
        key.SetValue("PersistBrowsers", 1, RegistryValueKind.DWord);
        Check.That(RegistryManager.RestoresFolderWindowsAtSignInUnder(root), "Folder Options stores the enabled option as 1.");
        key.SetValue("PersistBrowsers", 0, RegistryValueKind.DWord);
        Check.That(!RegistryManager.RestoresFolderWindowsAtSignInUnder(root), "0 turns the option off.");
        key.SetValue("PersistBrowsers", "1", RegistryValueKind.String);
        Check.That(!RegistryManager.RestoresFolderWindowsAtSignInUnder(root), "Only Explorer's DWORD value counts.");
    });

    private static Task MissingKeysStayMissingOnRead() => WithRoot(root =>
    {
        Check.That(!RegistryManager.IsStartupEnabledUnder(root), "A new profile must initially have startup disabled.");
        Check.Equal(0, root.SubKeyCount, "Reading startup state must not create Run or StartupApproved keys.");
    });

    private static Task MissingKeysAreCreatedOnEnable() => WithRoot(root =>
    {
        RegistryManager.ToggleStartup(root);
        Check.That(RegistryManager.IsStartupEnabledUnder(root), "First-time enabling must work without existing registry keys.");
        using (var run = root.OpenSubKey(RunPath))
        {
            Check.Equal($"\"{Helper.GetExecutablePath()}\" {Constants.BackgroundLaunchArg}", run?.GetValue(Constants.AppName) as string,
                "The startup command must quote the executable and launch in the background.");
        }
        using (var approval = root.OpenSubKey(ApprovedPath))
        {
            Check.That(approval?.GetValue(Constants.AppName) is byte[] { Length: 12 } data && data[0] == 2,
                "The approval entry must be enabled binary data.");
        }
        RegistryManager.ToggleStartup(root);
        Check.That(!RegistryManager.IsStartupEnabledUnder(root), "Disabling startup must remove the registration.");
        using var runAfter = root.OpenSubKey(RunPath);
        using var approvalAfter = root.OpenSubKey(ApprovedPath);
        Check.That(runAfter?.GetValue(Constants.AppName) == null && approvalAfter?.GetValue(Constants.AppName) == null,
            "Both startup values must be removed without deleting shared system keys.");
    });

    private static Task DisabledApprovalIsEnabled() => WithRoot(root =>
    {
        using (var run = root.CreateSubKey(RunPath))
            run.SetValue(Constants.AppName, $"\"{Helper.GetExecutablePath()}\" {Constants.BackgroundLaunchArg}");
        using (var approval = root.CreateSubKey(ApprovedPath))
            approval.SetValue(Constants.AppName, new byte[] { 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 }, RegistryValueKind.Binary);
        Check.That(!RegistryManager.IsStartupEnabledUnder(root), "An odd approval byte must keep startup disabled.");
        RegistryManager.ToggleStartup(root);
        Check.That(RegistryManager.IsStartupEnabledUnder(root), "Enabling must replace the disabled approval state.");
    });

    private static Task WithRoot(Action<RegistryKey> test)
    {
        // All operations are rooted under this unique test key, never the user's actual startup keys.
        var path = @"Software\WinTab.Tests\" + Guid.NewGuid().ToString("N");
        try
        {
            using var root = Registry.CurrentUser.CreateSubKey(path);
            test(root);
        }
        finally
        {
            Registry.CurrentUser.DeleteSubKeyTree(path, throwOnMissingSubKey: false);
        }
        return Task.CompletedTask;
    }
}
