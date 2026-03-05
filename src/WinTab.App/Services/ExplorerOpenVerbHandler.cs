using System;
using System.IO;
using System.Threading;
using WinTab.Diagnostics;
using WinTab.Platform.Win32;

namespace WinTab.App.Services;

public static class ExplorerOpenVerbHandler
{
    internal static Func<string, nint, bool> SendOpenFolderRequest =
        static (path, foreground) => ExplorerOpenRequestClient.TrySendOpenFolderEx(path, foreground);

    internal static Action<int> DelayBetweenRetries = Thread.Sleep;

    internal static Action<string, Logger?> OpenFolderFallback =
        static (path, logger) => AppEnvironment.TryOpenFolderFallback(path, logger);

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

        if (!AppEnvironment.TryNormalizeExistingDirectoryPath(args[1], out string path, out string reason))
        {
            effectiveLogger?.Warn($"Rejected open-folder invocation due to invalid path ({reason}). Raw='{args[1]}'");
            return true;
        }

        // Capture foreground BEFORE granting foreground rights - this is the accurate snapshot
        // of what was foreground at the moment the user initiated the open action.
        nint clickTimeForeground = (nint)NativeMethods.GetForegroundWindow();

        // Keep direct forwarding path to existing instance for smooth tab reuse.
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

        bool sent = false;
        for (int attempt = 0; attempt < 3 && !sent; attempt++)
        {
            sent = SendOpenFolderRequest(path, clickTimeForeground);
            if (!sent && attempt < 2)
                DelayBetweenRetries(80);
        }

        if (sent)
        {
            effectiveLogger?.Info($"Forwarded open-folder request to existing instance: {path}");
            return true;
        }

        effectiveLogger?.Warn($"No existing instance pipe; falling back to Explorer open without changing interception state: {path}");
        OpenFolderFallback(path, effectiveLogger);

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
