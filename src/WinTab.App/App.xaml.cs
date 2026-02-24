using System.Reflection;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using Microsoft.Extensions.DependencyInjection;
using Application = System.Windows.Application;
using WinTab.App.Services;
using WinTab.App.ExplorerTabUtilityPort;
using WinTab.App.ViewModels;
using WinTab.App.Views;
using WinTab.App.Views.Pages;
using WinTab.Core.Interfaces;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.Platform.Win32;
using WinTab.UI.Localization;
using WinTab.UI.Themes;

namespace WinTab.App;

public partial class App : Application
{
    private static Mutex? _singleInstanceMutex;
    private static bool _ownsSingleInstanceMutex;
    private static bool _explicitShutdownRequested;
    private const string ActivationEventName = "WinTab_ActivateMainWindow";
    private const string CleanupArgument = "--wintab-cleanup";
    private IServiceProvider? _serviceProvider;
    private TrayIconController? _trayIconController;
    private Logger? _logger;
    private AppLifecycleService? _lifecycleService;
    private ExplorerTabHookService? _explorerTabHook;
    private ExplorerTabMouseHookService? _explorerTabMouseHook;
    private ExplorerOpenRequestServer? _openRequestServer;
    private EventWaitHandle? _activationEvent;
    private CancellationTokenSource? _activationListenerCts;

    public static IServiceProvider Services { get; private set; } = null!;
    internal static bool IsExplicitShutdownRequested => _explicitShutdownRequested;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // -- 1. Install crash reporter ------------------------------------
        CrashReporter.Install(AppPaths.CrashLogPath);

        // -- 2. Command-line utility modes --------------------------------
        // Handle command modes before single-instance gate and before main logger lock.
        if (e.Args.Length >= 1 && string.Equals(e.Args[0], CleanupArgument, StringComparison.OrdinalIgnoreCase))
        {
            int code = RunUninstallCleanup();
            Shutdown(code);
            return;
        }

        if (e.Args.Length >= 1 && string.Equals(e.Args[0], "--wintab-companion", StringComparison.OrdinalIgnoreCase))
        {
            using Logger? companionLogger = TryCreateCompanionLogger();
            int code = ExplorerOpenVerbCompanion.Run(e.Args, companionLogger);
            Shutdown(code);
            return;
        }

        // When invoked via Explorer open-verb, forward the request to an existing instance.
        // If forwarding is unavailable, fallback to native Explorer open.
        if (TryHandleOpenFolderInvocation(e.Args, logger: null))
        {
            Shutdown();
            return;
        }

        // -- 3. Single instance check -------------------------------------
        _singleInstanceMutex = new Mutex(true, "WinTab_SingleInstance", out bool isNewInstance);
        _ownsSingleInstanceMutex = isNewInstance;
        if (!isNewInstance)
        {
            SignalExistingInstanceActivation();
            ActivateExistingInstance();
            Shutdown();
            return;
        }

        _activationEvent = new EventWaitHandle(false, EventResetMode.AutoReset, ActivationEventName);
        StartActivationListener();

        // -- 4. Create logger ---------------------------------------------
        _logger = new Logger(AppPaths.LogPath);
        _logger.Info("WinTab starting up...");

        // -- 5. Load settings ---------------------------------------------
        var settingsStore = new SettingsStore(AppPaths.SettingsPath, _logger);
        AppSettings settings = settingsStore.Load();

        // ThemeMode.System is no longer supported.
        // If an older settings file still has it, normalize to Light.
        if (settings.Theme == WinTab.Core.Enums.ThemeMode.System)
        {
            settings.Theme = WinTab.Core.Enums.ThemeMode.Light;
            settingsStore.Save(settings);
        }

        // -- 5.1 Pre-flight ------------------------------------------------
        // No-op placeholder: keep settings load spot for future migrations.

        // -- 6. Apply language --------------------------------------------
        LocalizationManager.ApplyLanguage(settings.Language);

        // -- 7. Apply theme -----------------------------------------------
        ThemeManager.ApplyTheme(settings.Theme);

        // -- 8. Configure DI ----------------------------------------------
        var services = new ServiceCollection();
        ConfigureServices(services, settings, settingsStore, _logger!);
        _serviceProvider = services.BuildServiceProvider();
        Services = _serviceProvider;

        // -- 9. Create and show MainWindow --------------------------------
        var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();

        if (settings.StartMinimized)
        {
            mainWindow.WindowState = WindowState.Minimized;
            mainWindow.ShowInTaskbar = false;
            // Do not call Show() -- tray icon is the access point.
        }
        else
        {
            mainWindow.Show();
        }

        MainWindow = mainWindow;

        // Re-apply once after window creation to ensure first paint uses
        // the configured theme resources on all visual elements.
        ThemeManager.ApplyTheme(settings.Theme);

        // -- 10. Start lifecycle services ---------------------------------
        _lifecycleService = _serviceProvider.GetRequiredService<AppLifecycleService>();
        _lifecycleService.Start();

        // Explorer integration (native Explorer tab conversion path).
        _explorerTabHook = _serviceProvider.GetRequiredService<ExplorerTabHookService>();
        _explorerTabMouseHook = _serviceProvider.GetRequiredService<ExplorerTabMouseHookService>();

        // IPC handler: allow handler invocations to forward open-folder requests.
        _openRequestServer = _serviceProvider.GetRequiredService<ExplorerOpenRequestServer>();
        _openRequestServer.Start(async path =>
        {
            try
            {
                var hook = _serviceProvider.GetRequiredService<ExplorerTabHookService>();
                bool handled = await hook.OpenLocationAsTabAsync(path);
                if (!handled)
                {
                    _logger?.Warn($"Open-folder request was not completed by tab hook; fallback open: {path}");
                    TryOpenFolderFallback(path, _logger);
                }
            }
            catch (Exception ex)
            {
                _logger?.Error("Failed to handle open-folder request.", ex);
                TryOpenFolderFallback(path, _logger);
            }
        });

        // Registry interception + companion (Win11 only).
        try
        {
            var interceptor = _serviceProvider.GetRequiredService<RegistryOpenVerbInterceptor>();

            bool isWin11 = OperatingSystem.IsWindowsVersionAtLeast(10, 0, 22000);
            string openVerbHandlerPath = ResolveLaunchExecutablePath();
            bool hasStableOpenVerbHandlerPath = IsStableOpenVerbHandlerPath(openVerbHandlerPath);
            bool enableExplorerOpenVerbInterception =
                settings.EnableExplorerOpenVerbInterception &&
                isWin11 &&
                hasStableOpenVerbHandlerPath;

            if (isWin11 && !hasStableOpenVerbHandlerPath)
            {
                _logger?.Warn($"Explorer open-verb interception disabled for transient executable path: {openVerbHandlerPath}");
            }

            // Always self-check first (repairs old crash residue).
            interceptor.StartupSelfCheck(settingEnabled: enableExplorerOpenVerbInterception);

            if (enableExplorerOpenVerbInterception)
            {
                interceptor.EnableOrRepair();
                StartCompanionWatcher();
            }
            else
            {
                // Ensure we do not leave overrides enabled on unsupported OS.
                interceptor.DisableAndRestore();
            }
        }
        catch (Exception ex)
        {
            _logger?.Error("Failed to configure Explorer open-verb interception.", ex);
        }

        // -- 11. Initialize tray icon -------------------------------------
        SetTrayIconVisibilityCore(true);

        _logger?.Info("WinTab started successfully.");
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _logger?.Info("WinTab shutting down...");

        try
        {
            _lifecycleService?.Stop();
            _trayIconController?.Dispose();
            _openRequestServer?.Dispose();
            _activationListenerCts?.Cancel();
            _activationListenerCts?.Dispose();
            _activationEvent?.Dispose();

            try
            {
                // Best-effort restore on clean exit.
                var interceptor = _serviceProvider?.GetService<RegistryOpenVerbInterceptor>();
                interceptor?.DisableAndRestore();
            }
            catch (Exception ex)
            {
                _logger?.Error("Failed to restore Explorer open-verb interception on exit.", ex);
            }

            if (_serviceProvider is IDisposable disposableProvider)
                disposableProvider.Dispose();

            _logger?.Info("WinTab shut down cleanly.");
            _logger?.Dispose();
        }
        catch (Exception ex)
        {
            _logger?.Error("Error during shutdown.", ex);
        }
        finally
        {
            // Release only if we successfully acquired initial ownership.
            if (_ownsSingleInstanceMutex)
            {
                try { _singleInstanceMutex?.ReleaseMutex(); }
                catch { /* ignore */ }
            }

            _singleInstanceMutex?.Dispose();
        }

        base.OnExit(e);
    }

    private static void ConfigureServices(
        IServiceCollection services,
        AppSettings settings,
        SettingsStore settingsStore,
        Logger logger)
    {
        services.AddSingleton(logger);
        services.AddSingleton(settingsStore);
        services.AddSingleton(settings);

        string exePath = ResolveLaunchExecutablePath();
        var startupRegistrar = new StartupRegistrar("WinTab", exePath);
        services.AddSingleton(startupRegistrar);

        services.AddSingleton<IWindowManager, WindowManager>();
        services.AddSingleton<IWindowEventSource, WindowEventWatcher>();

        services.AddSingleton<AppLifecycleService>();

        services.AddSingleton(sp =>
            new RegistryOpenVerbInterceptor(
                exePath,
                sp.GetRequiredService<Logger>()));

        services.AddSingleton<ExplorerOpenRequestServer>();

        // Back to native Explorer-tab pipeline (not overlay hijack).
        services.AddSingleton<ExplorerTabHookService>();
        services.AddSingleton<ExplorerTabMouseHookService>();

        services.AddSingleton<TrayIconController>(sp =>
            new TrayIconController(
                showSettings: () =>
                {
                    var window = Current.MainWindow;
                    if (window is not null)
                    {
                        window.Show();
                        window.WindowState = WindowState.Normal;
                        window.ShowInTaskbar = true;
                        
                        // Center window on screen when restored from tray
                        window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                        
                        window.Activate();
                    }
                },
                exitApp: RequestExplicitShutdown));

        services.AddTransient<GeneralViewModel>();
        services.AddTransient<BehaviorViewModel>();
        services.AddTransient<UninstallViewModel>();
        services.AddTransient<AboutViewModel>();

        services.AddSingleton<MainWindow>();
        services.AddTransient<GeneralPage>();
        services.AddTransient<BehaviorPage>();
        services.AddTransient<UninstallPage>();
        services.AddTransient<AboutPage>();
    }

    private static void ActivateExistingInstance()
    {
        try
        {
            using var current = System.Diagnostics.Process.GetCurrentProcess();
            var existing = System.Diagnostics.Process.GetProcessesByName(current.ProcessName)
                .Where(p => p.Id != current.Id)
                .OrderBy(p => p.StartTime)
                .FirstOrDefault(p => p.MainWindowHandle != IntPtr.Zero);

            if (existing is null)
                return;

            IntPtr hWnd = existing.MainWindowHandle;
            if (hWnd == IntPtr.Zero)
                return;

            NativeMethods.ShowWindow(hWnd, NativeConstants.SW_RESTORE);
            NativeMethods.SetForegroundWindow(hWnd);
        }
        catch
        {
            // Best effort only.
        }
    }

    private static void SignalExistingInstanceActivation()
    {
        try
        {
            using var activationEvent = EventWaitHandle.OpenExisting(ActivationEventName);
            activationEvent.Set();
        }
        catch
        {
        }
    }

    private void StartActivationListener()
    {
        if (_activationEvent is null)
            return;

        _activationListenerCts = new CancellationTokenSource();
        var token = _activationListenerCts.Token;
        var activationEvent = _activationEvent;

        Task.Run(() =>
        {
            while (!token.IsCancellationRequested)
            {
                activationEvent.WaitOne();
                if (token.IsCancellationRequested)
                    break;

                try
                {
                    Dispatcher.Invoke(() =>
                    {
                        if (Current.MainWindow is not Window window)
                            return;

                        if (!window.IsVisible)
                            window.Show();

                        window.ShowInTaskbar = true;
                        window.WindowState = WindowState.Normal;
                        window.Activate();
                    });
                }
                catch
                {
                }
            }
        }, token);
    }

    private static bool TryHandleOpenFolderInvocation(string[] args, Logger? logger)
    {
        using Logger? tempLogger = logger is null ? TryCreateCompanionLogger() : null;
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

        bool sent = ExplorerOpenRequestClient.TrySendOpenFolder(path);
        if (sent)
        {
            effectiveLogger?.Info($"Forwarded open-folder request to existing instance: {path}");
            return true;
        }

        effectiveLogger?.Warn($"No existing instance pipe; restoring shell defaults and falling back to Explorer open: {path}");
        TryRestoreExplorerOpenVerbDefaults(effectiveLogger);
        TryOpenFolderFallback(path, effectiveLogger);

        return true;
    }

    private static void TryRestoreExplorerOpenVerbDefaults(Logger? logger)
    {
        try
        {
            using RegistryKey? classesRoot = Registry.CurrentUser.OpenSubKey(@"Software\Classes", writable: true);
            if (classesRoot is not null)
            {
                classesRoot.DeleteSubKeyTree(@"Folder\shell\open\command", throwOnMissingSubKey: false);
                classesRoot.DeleteSubKeyTree(@"Directory\shell\open\command", throwOnMissingSubKey: false);
                classesRoot.DeleteSubKeyTree(@"Drive\shell\open\command", throwOnMissingSubKey: false);
            }

            using RegistryKey? folderShell = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Folder\shell", writable: true);
            using RegistryKey? directoryShell = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Directory\shell", writable: true);
            using RegistryKey? driveShell = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Drive\shell", writable: true);

            folderShell?.SetValue(string.Empty, "open", RegistryValueKind.String);
            directoryShell?.SetValue(string.Empty, "none", RegistryValueKind.String);
            driveShell?.SetValue(string.Empty, "none", RegistryValueKind.String);

            logger?.Info("Restored Explorer open-verb defaults for standalone handler invocation.");
        }
        catch (Exception ex)
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

    private static void TryOpenFolderFallback(string path, Logger? logger)
    {
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{path}\"",
                UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            logger?.Error($"Failed to fallback open-folder launch for path: {path}", ex);
        }
    }

    private static Logger? TryCreateCompanionLogger()
    {
        try
        {
            return new Logger(System.IO.Path.Combine(AppPaths.LogsDirectory, "wintab-companion.log"));
        }
        catch
        {
            return null;
        }
    }

    internal static void RequestExplicitShutdown()
    {
        if (_explicitShutdownRequested)
            return;

        _explicitShutdownRequested = true;
        Current.Shutdown();
    }

    private void SetTrayIconVisibilityCore(bool visible)
    {
        if (_serviceProvider is null)
            return;

        try
        {
            if (_trayIconController is null && !visible)
                return;

            _trayIconController ??= _serviceProvider.GetRequiredService<TrayIconController>();
            _trayIconController.SetVisible(visible);
        }
        catch (Exception ex)
        {
            _logger?.Error("Failed to update tray icon visibility.", ex);
        }
    }

    private void StartCompanionWatcher()
    {
        try
        {
            int pid = System.Diagnostics.Process.GetCurrentProcess().Id;
            string exePath = ResolveLaunchExecutablePath();

            // Companion is the same exe in a special mode.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = exePath,
                Arguments = $"--wintab-companion {pid}",
                UseShellExecute = false,
                CreateNoWindow = true,
            });
        }
        catch (Exception ex)
        {
            _logger?.Error("Failed to start companion watcher.", ex);
        }
    }

    private static string ResolveLaunchExecutablePath()
    {
        string? processPath = Environment.ProcessPath;
        if (!string.IsNullOrWhiteSpace(processPath) && !IsDotNetHost(processPath))
            return processPath;

        string assemblyPath = Assembly.GetExecutingAssembly().Location;
        if (!string.IsNullOrWhiteSpace(assemblyPath))
        {
            string appHostPath = Path.ChangeExtension(assemblyPath, ".exe");
            if (File.Exists(appHostPath))
                return appHostPath;
        }

        return processPath ?? assemblyPath;
    }

    private static bool IsDotNetHost(string path)
    {
        string fileName = Path.GetFileName(path);
        return string.Equals(fileName, "dotnet", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(fileName, "dotnet.exe", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsStableOpenVerbHandlerPath(string exePath)
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

    private static int RunUninstallCleanup()
    {
        string exePath = ResolveLaunchExecutablePath();
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
            catch (Exception ex)
            {
                failureCount++;
                cleanupLogger?.Error("Uninstall cleanup failed to restore Explorer open-verb state.", ex);
                TryRestoreExplorerOpenVerbDefaults(cleanupLogger);
            }

            TryDeleteExplorerOpenVerbBackupRegistryCache();

            cleanupLogger?.Info("Uninstall cleanup completed.");
            return failureCount == 0 ? 0 : 1;
        }
        finally
        {
            cleanupLogger?.Dispose();
        }
    }
}
