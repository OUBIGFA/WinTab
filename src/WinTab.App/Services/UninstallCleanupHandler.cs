using System;
using System.IO;
using Microsoft.Win32;
using WinTab.Diagnostics;
using WinTab.Platform.Win32;
using WinTab.Core.Models;
using WinTab.Persistence;

namespace WinTab.App.Services;

public static class UninstallCleanupHandler
{
    private const string DelegateExecuteClsidBraced = "{FD5BF2CD-0B24-4A80-9AF3-E40F9AFC0001}";
    public static int RunUninstallCleanup(string exePath)
    {
        Logger? cleanupLogger = null;
        int failureCount = 0;

        try
        {
            try
            {
                var startupRegistrar = new StartupRegistrar("WinTab", exePath);
                startupRegistrar.SetEnabled(false);

                if (startupRegistrar.IsEnabled())
                    failureCount++;
            }
            catch
            {
                failureCount++;
            }

            try
            {
                cleanupLogger = TryCreateCompanionLogger();
                if (cleanupLogger is null)
                {
                    string tempLogPath = Path.Combine(Path.GetTempPath(), "WinTab", "wintab-cleanup.log");
                    cleanupLogger = new Logger(tempLogPath);
                }

                var interceptor = new RegistryOpenVerbInterceptor(exePath, cleanupLogger);
                interceptor.DisableAndRestore();
            }
            catch (Exception ex) when (ex is UnauthorizedAccessException or System.Security.SecurityException or IOException)
            {
                failureCount++;
                cleanupLogger?.Error("Uninstall cleanup failed to restore Explorer open-verb state.", ex);
                TryRestoreExplorerOpenVerbDefaults(cleanupLogger);
            }

            TryDeleteExplorerOpenVerbBackupRegistryCache();
            TryDeleteWinTabRegistryTree();

            cleanupLogger?.Info("Uninstall cleanup completed.");
            return failureCount == 0 ? 0 : 1;
        }
        finally
        {
            cleanupLogger?.Dispose();
        }
    }

    public static void TryRestoreExplorerOpenVerbDefaults(Logger? logger)
    {
        try
        {
            using RegistryKey? classesRoot = Registry.CurrentUser.OpenSubKey(@"Software\Classes", writable: true);
            if (classesRoot is not null)
            {
                string[] classes = new[] { "Folder", "Directory", "Drive" };
                string[] verbs = new[] { "open", "explore", "opennewwindow" };

                foreach (string cls in classes)
                {
                    using RegistryKey? shell = Registry.CurrentUser.OpenSubKey($@"Software\Classes\{cls}\shell", writable: true);
                    shell?.DeleteValue(string.Empty, throwOnMissingValue: false);

                    foreach (string verb in verbs)
                        classesRoot.DeleteSubKeyTree($@"{cls}\shell\{verb}", throwOnMissingSubKey: false);

                    TryDeleteEmptyKey(classesRoot, $@"{cls}\shell");
                    TryDeleteEmptyKey(classesRoot, cls);
                }
            }

            // Clean up HKCU COM registration (user-only)
            foreach (RegistryView view in GetRegistryViews())
            {
                using RegistryKey? clsidRoot = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view)
                    .OpenSubKey(@"Software\Classes\CLSID", writable: true);
                clsidRoot?.DeleteSubKeyTree(DelegateExecuteClsidBraced, throwOnMissingSubKey: false);
            }

            // Clean up HKLM COM registration (machine-wide) - added for Windows 11 Start Menu support
            foreach (RegistryView view in GetRegistryViews())
            {
                try
                {
                    using RegistryKey? clsidRoot = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view)
                        .OpenSubKey(@"Software\Classes\CLSID", writable: true);
                    clsidRoot?.DeleteSubKeyTree(DelegateExecuteClsidBraced, throwOnMissingSubKey: false);
                }
                catch (Exception ex) when (ex is UnauthorizedAccessException or System.Security.SecurityException)
                {
                    // HKLM cleanup may fail without admin rights - this is expected during uninstall
                    logger?.Warn($"Failed to clean up HKLM COM registration (view {view}): {ex.Message}");
                }
            }

            logger?.Info("Removed WinTab overrides and restored native Explorer defaults for standalone cleanup.");
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or System.Security.SecurityException or IOException)
        {
            logger?.Error("Failed to restore Explorer open-verb defaults.", ex);
        }
    }

    private static void TryDeleteEmptyKey(RegistryKey root, string subKeyPath)
    {
        using RegistryKey? key = root.OpenSubKey(subKeyPath, writable: true);
        if (key is null)
            return;

        if (key.SubKeyCount == 0 && key.ValueCount == 0)
            root.DeleteSubKeyTree(subKeyPath, throwOnMissingSubKey: false);
    }

    private static RegistryView[] GetRegistryViews()
    {
        return Environment.Is64BitOperatingSystem
            ? [RegistryView.Registry64, RegistryView.Registry32]
            : [RegistryView.Registry32];
    }

    private static void TryDeleteExplorerOpenVerbBackupRegistryCache()
    {
        try
        {
            using RegistryKey? softwareRoot = Registry.CurrentUser.OpenSubKey(@"Software", writable: true);
            softwareRoot?.DeleteSubKeyTree(@"WinTab\Backups\ExplorerOpenVerb", throwOnMissingSubKey: false);
        }
        catch
        {
        }
    }

    public static void TryDeleteWinTabRegistryTree()
    {
        try
        {
            using RegistryKey? softwareRoot = Registry.CurrentUser.OpenSubKey(@"Software", writable: true);
            softwareRoot?.DeleteSubKeyTree(@"WinTab", throwOnMissingSubKey: false);
        }
        catch
        {
        }
    }

    public static Logger? TryCreateCompanionLogger()
    {
        try
        {
            return new Logger(Path.Combine(AppPaths.LogsDirectory, "wintab-companion.log"));
        }
        catch
        {
            return null;
        }
    }
}
