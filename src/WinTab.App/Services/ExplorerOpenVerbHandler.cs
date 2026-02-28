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

        string path = args[1].Trim().Trim('"');
        if (string.IsNullOrWhiteSpace(path))
            return true;

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

        effectiveLogger?.Warn($"No existing instance pipe; falling back to Explorer open without changing interception state: {path}");
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
