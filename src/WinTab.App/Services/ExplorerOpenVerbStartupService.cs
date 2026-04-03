using System.Threading;
using WinTab.Core.Models;
using WinTab.Diagnostics;

namespace WinTab.App.Services;

public sealed class ExplorerOpenVerbStartupService
    : IExplorerOpenVerbConfigurationController
{
    private readonly IExplorerOpenVerbInterceptor _interceptor;
    private readonly Logger _logger;
    private readonly Func<bool> _isWindows11;
    private readonly Func<string> _resolveLaunchExecutablePath;
    private readonly Func<string, bool> _isStableOpenVerbHandlerPath;
    private readonly Func<Func<Task>, Task> _runInBackground;

    private int _started;
    private Task? _startupTask;
    private readonly SemaphoreSlim _configurationGate = new(1, 1);

    public ExplorerOpenVerbStartupService(
        IExplorerOpenVerbInterceptor interceptor,
        Logger logger,
        WinTab.Platform.Win32.StartupRegistrar startupRegistrar)
        : this(
            interceptor,
            logger,
            isWindows11: () => OperatingSystem.IsWindowsVersionAtLeast(10, 0, 22000),
            resolveLaunchExecutablePath: AppEnvironment.ResolveLaunchExecutablePath,
            isStableOpenVerbHandlerPath: ExplorerOpenVerbHandler.IsStableOpenVerbHandlerPath,
            runInBackground: work => Task.Run(work))
    {
    }

    public ExplorerOpenVerbStartupService(
        IExplorerOpenVerbInterceptor interceptor,
        Logger logger,
        Func<bool> isWindows11,
        Func<string> resolveLaunchExecutablePath,
        Func<string, bool> isStableOpenVerbHandlerPath,
        Func<Func<Task>, Task> runInBackground)
    {
        ArgumentNullException.ThrowIfNull(interceptor);
        ArgumentNullException.ThrowIfNull(logger);
        ArgumentNullException.ThrowIfNull(isWindows11);
        ArgumentNullException.ThrowIfNull(resolveLaunchExecutablePath);
        ArgumentNullException.ThrowIfNull(isStableOpenVerbHandlerPath);
        ArgumentNullException.ThrowIfNull(runInBackground);

        _interceptor = interceptor;
        _logger = logger;
        _isWindows11 = isWindows11;
        _resolveLaunchExecutablePath = resolveLaunchExecutablePath;
        _isStableOpenVerbHandlerPath = isStableOpenVerbHandlerPath;
        _runInBackground = runInBackground;
    }

    public void Start(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        if (Interlocked.Exchange(ref _started, 1) == 1)
            return;

        AppSettings snapshot = CreateStartupSnapshot(settings);
        if (ExplorerOpenVerbInterceptionPolicy.NormalizeForNativeCurrentDirectoryBehavior(snapshot))
            settings.EnableExplorerOpenVerbInterception = snapshot.EnableExplorerOpenVerbInterception;

        try
        {
            // Arm or disarm interception before startup continues so the app
            // never reports itself ready while Explorer is still on the old
            // flash-then-merge path.
            _startupTask = ConfigureSerializedAsync(snapshot);
            _startupTask.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to queue Explorer open-verb startup configuration.", ex);
        }
    }

    public void ReconfigureForCurrentSettings(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        AppSettings snapshot = CreateStartupSnapshot(settings);
        if (ExplorerOpenVerbInterceptionPolicy.NormalizeForNativeCurrentDirectoryBehavior(snapshot))
            settings.EnableExplorerOpenVerbInterception = snapshot.EnableExplorerOpenVerbInterception;

        try
        {
            Task task = ConfigureSerializedAsync(snapshot);
            _startupTask = task;
            task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to reconfigure Explorer open-verb interception for runtime settings change.", ex);
        }
    }

    public void WaitForStartupConfigurationToFinish(TimeSpan timeout)
    {
        Task? task = _startupTask;
        if (task is null || task.IsCompleted)
            return;

        try
        {
            if (!task.Wait(timeout))
            {
                _logger.Warn("Timed out while waiting for Explorer open-verb startup configuration to complete.");
            }
        }
        catch (Exception ex)
        {
            _logger.Error("Explorer open-verb startup configuration failed before shutdown.", ex);
        }
    }

    private static AppSettings CreateStartupSnapshot(AppSettings source)
    {
        return new AppSettings
        {
            EnableExplorerOpenVerbInterception = source.EnableExplorerOpenVerbInterception,
            OpenChildFolderInNewTabFromActiveTab = source.OpenChildFolderInNewTabFromActiveTab,
            EnableAutoConvertExplorerWindows = source.EnableAutoConvertExplorerWindows,
            PersistExplorerOpenVerbInterceptionAcrossExit = source.PersistExplorerOpenVerbInterceptionAcrossExit,
            RunAtStartup = source.RunAtStartup,
        };
    }

    private async Task ConfigureSerializedAsync(AppSettings settings)
    {
        await _configurationGate.WaitAsync().ConfigureAwait(false);

        try
        {
            ExplorerOpenVerbInterceptionPolicy.NormalizeForNativeCurrentDirectoryBehavior(settings);

            string openVerbHandlerPath = _resolveLaunchExecutablePath();
            bool hasStableOpenVerbHandlerPath = _isStableOpenVerbHandlerPath(openVerbHandlerPath);
            bool isWin11 = _isWindows11();
            bool enableExplorerOpenVerbInterception =
                ExplorerOpenVerbInterceptionPolicy.ShouldEnableOpenVerbInterception(settings, hasStableOpenVerbHandlerPath);
            bool persistAcrossReboot = ExplorerOpenVerbInterceptionPolicy.ShouldPersistAcrossReboot(settings);

            if (!hasStableOpenVerbHandlerPath)
            {
                _logger.Warn($"Explorer open-verb interception disabled for transient executable path: {openVerbHandlerPath}");
            }

            if (!isWin11 && settings.EnableExplorerOpenVerbInterception)
            {
                _logger.Warn("Explorer open-verb interception is running in compatibility mode on non-Windows 11 systems.");
            }

            _interceptor.StartupSelfCheck(
                settingEnabled: enableExplorerOpenVerbInterception,
                persistAcrossReboot: persistAcrossReboot);

            if (enableExplorerOpenVerbInterception)
            {
                _interceptor.EnableOrRepair(persistAcrossReboot);
            }
            else
            {
                _interceptor.DisableAndRestore(deleteBackup: false);
            }
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to configure Explorer open-verb interception.", ex);
        }
        finally
        {
            _configurationGate.Release();
        }
    }
}
