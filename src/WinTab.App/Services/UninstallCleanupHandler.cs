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
                    foreach (string verb in verbs)
                    {
                        classesRoot.DeleteSubKeyTree($@"{cls}\shell\{verb}\command", throwOnMissingSubKey: false);
                    }
                }
            }

            using RegistryKey? folderShell = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Folder\shell", writable: true);
            using RegistryKey? directoryShell = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Directory\shell", writable: true);
            using RegistryKey? driveShell = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Drive\shell", writable: true);

            folderShell?.SetValue(string.Empty, "open", RegistryValueKind.String);
            directoryShell?.SetValue(string.Empty, "none", RegistryValueKind.String);
            driveShell?.SetValue(string.Empty, "none", RegistryValueKind.String);

            logger?.Info("Restored Explorer open-verb defaults for standalone handler invocation.");
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or System.Security.SecurityException or IOException)
        {
            logger?.Error("Failed to restore Explorer open-verb defaults.", ex);
        }
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
