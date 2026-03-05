using System.Threading;
using WinTab.Core.Models;
using WinTab.Diagnostics;

namespace WinTab.App.Services;

public sealed class ExplorerOpenVerbStartupService
{
    private readonly IExplorerOpenVerbInterceptor _interceptor;
    private readonly Logger _logger;
    private readonly Func<bool> _isWindows11;
    private readonly Func<string> _resolveLaunchExecutablePath;
    private readonly Func<string, bool> _isStableOpenVerbHandlerPath;
    private readonly Func<Func<Task>, Task> _runInBackground;

    private int _started;
    private Task? _startupTask;

    public ExplorerOpenVerbStartupService(
        IExplorerOpenVerbInterceptor interceptor,
        Logger logger)
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

        try
        {
            // If we need to DISABLE interception (e.g. preserve native in-Explorer browsing),
            // do it synchronously to avoid a race where Explorer still launches the handler
            // before the background task runs.
            string openVerbHandlerPath = _resolveLaunchExecutablePath();
            bool hasStableOpenVerbHandlerPath = _isStableOpenVerbHandlerPath(openVerbHandlerPath);
            bool enableExplorerOpenVerbInterception =
                ExplorerOpenVerbInterceptionPolicy.ShouldEnableOpenVerbInterception(snapshot, hasStableOpenVerbHandlerPath);

            if (!enableExplorerOpenVerbInterception)
            {
                _startupTask = ConfigureAsync(snapshot);
                return;
            }

            _startupTask = _runInBackground(() => ConfigureAsync(snapshot));
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to queue Explorer open-verb startup configuration.", ex);
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
        };
    }

    private Task ConfigureAsync(AppSettings settings)
    {
        try
        {
            string openVerbHandlerPath = _resolveLaunchExecutablePath();
            bool hasStableOpenVerbHandlerPath = _isStableOpenVerbHandlerPath(openVerbHandlerPath);
            bool isWin11 = _isWindows11();

            bool enableExplorerOpenVerbInterception =
                ExplorerOpenVerbInterceptionPolicy.ShouldEnableOpenVerbInterception(settings, hasStableOpenVerbHandlerPath);

            if (!hasStableOpenVerbHandlerPath)
            {
                _logger.Warn($"Explorer open-verb interception disabled for transient executable path: {openVerbHandlerPath}");
            }

            if (!isWin11 && settings.EnableExplorerOpenVerbInterception)
            {
                _logger.Warn("Explorer open-verb interception is running in compatibility mode on non-Windows 11 systems.");
            }

            _interceptor.StartupSelfCheck(settingEnabled: enableExplorerOpenVerbInterception);

            if (enableExplorerOpenVerbInterception)
            {
                _interceptor.EnableOrRepair();
            }
            else
            {
                _interceptor.DisableAndRestore();
            }
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to configure Explorer open-verb interception.", ex);
        }

        return Task.CompletedTask;
    }
}
