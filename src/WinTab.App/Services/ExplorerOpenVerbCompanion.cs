using System.Diagnostics;
using System.IO;
using System.Threading;
using WinTab.Diagnostics;
using WinTab.Persistence;

namespace WinTab.App.Services;

public static class ExplorerOpenVerbCompanion
{
    public static int Run(string[] args, Logger? logger)
    {
        Logger? ownedLogger = null;

        try
        {
            if (logger is null)
            {
                try
                {
                    ownedLogger = new Logger(Path.Combine(AppPaths.LogsDirectory, "wintab-companion.log"));
                    logger = ownedLogger;
                }
                catch
                {
                    return 1;
                }
            }

            if (logger is null)
                return 1;

            if (args.Length < 2 || !int.TryParse(args[1], out int pid))
            {
                logger?.Warn("Companion: missing parent PID.");
                return 2;
            }

            Process parent;
            try
            {
                parent = Process.GetProcessById(pid);
            }
            catch
            {
                logger?.Warn($"Companion: parent PID {pid} is not running.");
                return 0;
            }

            using (parent)
            {
                logger?.Info($"Companion: watching parent PID {pid}.");
                logger?.Info($"Companion: started for parent PID {pid}.");

                // Wait for parent exit.
                while (true)
                {
                    if (parent.HasExited)
                        break;

                    Thread.Sleep(500);
                }

                // If the parent exited cleanly, it should have restored and removed the override.
                // If not, restore now.
                // Note: do not rely on current exe path for detection; the registry check uses marker args.
                // But for restore we still want the original process exe path so we can write the override command correctly.
                string parentExePath = parent.MainModule?.FileName ?? "wintab";
                var interceptor = new RegistryOpenVerbInterceptor(exePath: parentExePath, logger!);

                if (interceptor.IsEnabled())
                {
                    logger?.Warn("Companion: parent exited but Explorer open-verb still points to WinTab; restoring.");
                    interceptor.DisableAndRestore();
                }

                logger?.Info("Companion: done.");
                return 0;
            }

            // Unreachable (kept empty intentionally).
        }
        catch (Exception ex)
        {
            logger?.Error("Companion: fatal error.", ex);
            return 1;
        }
        finally
        {
            ownedLogger?.Dispose();
        }
    }
}
