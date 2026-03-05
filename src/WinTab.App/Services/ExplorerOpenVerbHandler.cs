using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.Platform.Win32;

namespace WinTab.App.Services;

public static class ExplorerOpenVerbHandler
{
    internal static Func<string, nint, bool> SendOpenFolderRequest =
        static (path, foreground) => ExplorerOpenRequestClient.TrySendOpenFolderEx(path, foreground);

    internal static Action<int> DelayBetweenRetries = Thread.Sleep;

    internal static Action<string, Logger?> OpenFolderFallback =
        static (path, logger) => AppEnvironment.TryOpenFolderFallback(path, logger);

    internal static Func<nint, string, bool> ShouldBypassInterceptionAndUseNativeOpen =
        static (foreground, targetPath) => ShouldUseNativeCurrentDirectoryBehavior((IntPtr)foreground, targetPath);

    internal static Func<nint, bool> IsExplorerTopLevelWindowPredicate =
        static foreground => IsExplorerTopLevelWindow((IntPtr)foreground);

    internal static Func<bool> LoadOpenChildFolderInNewTabSettingPredicate =
        static () => LoadOpenChildFolderInNewTabSetting();

    internal static Func<nint, string?> TryGetForegroundExplorerDirectoryPredicate =
        static foreground => TryGetForegroundExplorerDirectory((IntPtr)foreground);

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

        // Capture foreground BEFORE granting foreground rights — this is the accurate snapshot
        // of what was foreground at the moment the user initiated the open action.
        nint clickTimeForeground = (nint)NativeMethods.GetForegroundWindow();

        if (ShouldBypassInterceptionAndUseNativeOpen(clickTimeForeground, path))
        {
            effectiveLogger?.Info($"Current-directory browse with child-folder-new-tab disabled; using native Explorer open directly: {path}");
            OpenFolderFallback(path, effectiveLogger);
            return true;
        }

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

    private static bool ShouldUseNativeCurrentDirectoryBehavior(IntPtr clickTimeForeground, string targetPath)
    {
        if (!IsExplorerTopLevelWindowPredicate((nint)clickTimeForeground))
            return false;

        if (LoadOpenChildFolderInNewTabSettingPredicate())
            return false;

        string? currentDirectory = TryGetForegroundExplorerDirectoryPredicate((nint)clickTimeForeground);
        if (string.IsNullOrWhiteSpace(currentDirectory))
            return false;

        return IsDirectChildPathOf(currentDirectory, targetPath);
    }

    private static bool LoadOpenChildFolderInNewTabSetting()
    {
        try
        {
            using var store = new SettingsStore(AppPaths.SettingsPath);
            return store.Load().OpenChildFolderInNewTabFromActiveTab;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsExplorerTopLevelWindow(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero || !NativeMethods.IsWindow(hwnd))
            return false;

        IntPtr topLevel = NativeMethods.GetAncestor(hwnd, NativeConstants.GA_ROOT);
        if (topLevel == IntPtr.Zero || !NativeMethods.IsWindow(topLevel))
            return false;

        var classBuilder = new StringBuilder(64);
        NativeMethods.GetClassName(topLevel, classBuilder, classBuilder.Capacity);
        if (!string.Equals(classBuilder.ToString(), "CabinetWClass", StringComparison.OrdinalIgnoreCase))
            return false;

        NativeMethods.GetWindowThreadProcessId(topLevel, out uint pid);
        if (pid == 0)
            return false;

        try
        {
            using Process process = Process.GetProcessById((int)pid);
            return string.Equals(process.ProcessName, "explorer", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static string? TryGetForegroundExplorerDirectory(IntPtr foregroundWindow)
    {
        try
        {
            IntPtr topLevel = NativeMethods.GetAncestor(foregroundWindow, NativeConstants.GA_ROOT);
            if (topLevel == IntPtr.Zero)
                return null;

            Type? shellAppType = Type.GetTypeFromProgID("Shell.Application");
            if (shellAppType is null)
                return null;

            object? shellApp = null;
            object? windows = null;

            try
            {
                shellApp = Activator.CreateInstance(shellAppType);
                if (shellApp is null)
                    return null;

                windows = shellAppType.InvokeMember("Windows", System.Reflection.BindingFlags.InvokeMethod, null, shellApp, null);
                if (windows is null)
                    return null;

                int count = (int)windows.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, windows, null)!;
                for (int i = 0; i < count; i++)
                {
                    object? window = null;
                    try
                    {
                        window = windows.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, windows, [i]);
                        if (window is null)
                            continue;

                        int hwnd = (int)window.GetType().InvokeMember("HWND", System.Reflection.BindingFlags.GetProperty, null, window, null)!;
                        if (new IntPtr(hwnd) != topLevel)
                            continue;

                        string? path = window.GetType().InvokeMember("Document", System.Reflection.BindingFlags.GetProperty, null, window, null)?
                            .GetType().InvokeMember("Folder", System.Reflection.BindingFlags.GetProperty, null, window.GetType().InvokeMember("Document", System.Reflection.BindingFlags.GetProperty, null, window, null), null)?
                            .GetType().InvokeMember("Self", System.Reflection.BindingFlags.GetProperty, null, window.GetType().InvokeMember("Document", System.Reflection.BindingFlags.GetProperty, null, window, null)?.GetType().InvokeMember("Folder", System.Reflection.BindingFlags.GetProperty, null, window.GetType().InvokeMember("Document", System.Reflection.BindingFlags.GetProperty, null, window, null), null), null)?
                            .GetType().InvokeMember("Path", System.Reflection.BindingFlags.GetProperty, null, window.GetType().InvokeMember("Document", System.Reflection.BindingFlags.GetProperty, null, window, null)?.GetType().InvokeMember("Folder", System.Reflection.BindingFlags.GetProperty, null, window.GetType().InvokeMember("Document", System.Reflection.BindingFlags.GetProperty, null, window, null), null)?.GetType().InvokeMember("Self", System.Reflection.BindingFlags.GetProperty, null, window.GetType().InvokeMember("Document", System.Reflection.BindingFlags.GetProperty, null, window, null)?.GetType().InvokeMember("Folder", System.Reflection.BindingFlags.GetProperty, null, window.GetType().InvokeMember("Document", System.Reflection.BindingFlags.GetProperty, null, window, null), null), null), null)?.ToString();

                        if (string.IsNullOrWhiteSpace(path))
                            continue;

                        return Path.GetFullPath(path);
                    }
                    catch
                    {
                        // skip unstable COM window
                    }
                    finally
                    {
                        if (window is not null && Marshal.IsComObject(window))
                            Marshal.FinalReleaseComObject(window);
                    }
                }
            }
            finally
            {
                if (windows is not null && Marshal.IsComObject(windows))
                    Marshal.FinalReleaseComObject(windows);
                if (shellApp is not null && Marshal.IsComObject(shellApp))
                    Marshal.FinalReleaseComObject(shellApp);
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    private static bool IsDirectChildPathOf(string currentDirectory, string targetPath)
    {
        try
        {
            string normalizedCurrent = NormalizePathForCompare(Path.GetFullPath(currentDirectory));
            string normalizedTarget = NormalizePathForCompare(Path.GetFullPath(targetPath));

            if (!normalizedTarget.StartsWith(normalizedCurrent, StringComparison.OrdinalIgnoreCase))
                return false;

            if (normalizedTarget.Length <= normalizedCurrent.Length)
                return false;

            return normalizedTarget[normalizedCurrent.Length] == '\\';
        }
        catch
        {
            return false;
        }
    }

    private static string NormalizePathForCompare(string path)
    {
        return path.TrimEnd('\\').Replace('/', '\\');
    }
}
