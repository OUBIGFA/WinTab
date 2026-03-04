using System;
using System.IO;
using WinTab.Diagnostics;
using WinTab.Platform.Win32;

namespace WinTab.App.Services;

public static class ExplorerOpenVerbHandler
{
    public static bool TryHandleOpenFolderInvocation(string[] args, Logger? logger)
    {
        using Logger? tempLogger = logger is null ? UninstallCleanupHandler.TryCreateCompanionLogger() : null;
        Logger? effectiveLogger = logger ?? tempLogger;

        // Registry handler: "WinTab.exe --wintab-open-folder \"%1\""
        if (args.Length < 2)
            return false;

        if (!string.Equals(args[0], RegistryOpenVerbInterceptor.HandlerArgument, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(args[0], "--open-folder", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        string rawPath = args[1];
        if (!AppEnvironment.TryNormalizeExistingDirectoryPath(rawPath, out string path, out string reason))
        {
            effectiveLogger?.Warn($"Rejected open-folder invocation due to invalid path ({reason}). Raw='{rawPath}'");
            return true;
        }

        // Capture foreground BEFORE granting foreground rights — this is the accurate snapshot
        // of what was foreground at the moment the user initiated the open action.
        nint clickTimeForeground = (nint)NativeMethods.GetForegroundWindow();

        try
        {
            // The handler process is launched by a user-initiated shell action,
            // so grant foreground rights to improve focus handoff to the existing instance.
            NativeMethods.AllowSetForegroundWindow(NativeConstants.ASFW_ANY);
        }
        catch
        {
            // best effort
        }

        bool sent = ExplorerOpenRequestClient.TrySendOpenFolderEx(path, clickTimeForeground);
        if (sent)
        {
            effectiveLogger?.Info($"Forwarded open-folder request to existing instance: {path}");
            return true;
        }

        Logger? transientRecoveryLogger = null;
        try
        {
            string exePath = AppEnvironment.ResolveLaunchExecutablePath();
            Logger? recoveryLogger = effectiveLogger;
            if (recoveryLogger is null)
            {
                string tempLogPath = Path.Combine(Path.GetTempPath(), "WinTab", "wintab-handler-recovery.log");
                recoveryLogger = new Logger(tempLogPath);
                transientRecoveryLogger = recoveryLogger;
            }

            var interceptor = new RegistryOpenVerbInterceptor(exePath, recoveryLogger);
            interceptor.DisableAndRestore();
            effectiveLogger?.Info($"No existing WinTab pipe available; restored native Explorer open-verb defaults before fallback: {path}");
        }
        catch (Exception ex)
        {
            effectiveLogger?.Error("Failed to restore native Explorer open-verb defaults after pipe forwarding failure.", ex);
        }
        finally
        {
            transientRecoveryLogger?.Dispose();
        }

        effectiveLogger?.Warn($"No existing instance pipe; falling back to native Explorer open: {path}");
        AppEnvironment.TryOpenFolderFallback(path, effectiveLogger);

        return true;
    }

    public static bool IsStableOpenVerbHandlerPath(string exePath)
    {
        if (string.IsNullOrWhiteSpace(exePath))
            return false;

        try
        {
            string fullPath = Path.GetFullPath(exePath).Replace('/', '\\');
            if (!File.Exists(fullPath))
                return false;

            if (fullPath.Contains("\\tasks\\build_tmp\\", StringComparison.OrdinalIgnoreCase))
                return false;

            if (fullPath.Contains("\\obj\\", StringComparison.OrdinalIgnoreCase))
                return false;

            return true;
        }
        catch
        {
            return false;
        }
    }
}
