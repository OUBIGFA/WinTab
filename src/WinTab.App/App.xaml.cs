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
    private static bool _explicitShutdownRequested;
    private const string CleanupArgument = "--wintab-cleanup";
    private IServiceProvider? _serviceProvider;
    private TrayIconController? _trayIconController;
    private Logger? _logger;
    private SingleInstanceService? _singleInstanceService;
    private AppLifecycleService? _lifecycleService;
    private ExplorerTabHookService? _explorerTabHook;
    private ExplorerTabMouseHookService? _explorerTabMouseHook;
    private ExplorerOpenRequestServer? _openRequestServer;

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
            int code = UninstallCleanupHandler.RunUninstallCleanup(AppEnvironment.ResolveLaunchExecutablePath());
            Shutdown(code);
            return;
        }

        if (e.Args.Length >= 1 && string.Equals(e.Args[0], "--wintab-companion", StringComparison.OrdinalIgnoreCase))
        {
            Shutdown(0);
            return;
        }

        // When invoked via Explorer open-verb, forward the request to an existing instance.
        // If forwarding is unavailable, fallback to native Explorer open.
        if (ExplorerOpenVerbHandler.TryHandleOpenFolderInvocation(e.Args, logger: null))
        {
            Shutdown();
            return;
        }

        // -- 3. Single instance check -------------------------------------
        _singleInstanceService = new SingleInstanceService();
        if (!_singleInstanceService.InitializeAsFirstInstance())
        {
            _singleInstanceService.SignalExistingInstanceActivation();
            _singleInstanceService.BringExistingInstanceToForeground();
            Shutdown();
            return;
        }

        _singleInstanceService.StartActivationListener();

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
        _openRequestServer.Start(async (path, clickTimeForeground) =>
        {
            try
            {
                var hook = _serviceProvider.GetRequiredService<ExplorerTabHookService>();
                bool handled = await hook.OpenInterceptedLocationAsTabAsync(path, clickTimeForeground);
                if (!handled)
                {
                    _logger?.Warn($"[Intercept] Open-folder request handling failed: {path}");
                }
            }
            catch (Exception ex)
            {
                _logger?.Error("Failed to handle open-folder request.", ex);
            }
        });

        // Registry interception + companion (Win11 only).
        try
        {
            var interceptor = _serviceProvider.GetRequiredService<RegistryOpenVerbInterceptor>();

            bool isWin11 = OperatingSystem.IsWindowsVersionAtLeast(10, 0, 22000);
            string openVerbHandlerPath = AppEnvironment.ResolveLaunchExecutablePath();
            bool hasStableOpenVerbHandlerPath = ExplorerOpenVerbHandler.IsStableOpenVerbHandlerPath(openVerbHandlerPath);
            bool enableExplorerOpenVerbInterception =
                settings.EnableExplorerOpenVerbInterception &&
                hasStableOpenVerbHandlerPath;

            if (!hasStableOpenVerbHandlerPath)
            {
                _logger?.Warn($"Explorer open-verb interception disabled for transient executable path: {openVerbHandlerPath}");
            }

            if (!isWin11 && settings.EnableExplorerOpenVerbInterception)
            {
                _logger?.Warn("Explorer open-verb interception is running in compatibility mode on non-Windows 11 systems.");
            }

            // Always self-check first (repairs old crash residue).
            interceptor.StartupSelfCheck(settingEnabled: enableExplorerOpenVerbInterception);

            if (enableExplorerOpenVerbInterception)
            {
                interceptor.EnableOrRepair();
            }
            else
            {
                interceptor.DisableAndRestore();
            }
        }
        catch (Exception ex)
        {
            _logger?.Error("Failed to configure Explorer open-verb interception.", ex);
        }

        // -- 11. Initialize tray icon -------------------------------------
        SetTrayIconVisibility(settings.ShowTrayIcon);

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

            try
            {
                AppSettings? settings = _serviceProvider?.GetService<AppSettings>();
                if (settings is not null &&
                    settings.EnableExplorerOpenVerbInterception)
                {
                    var interceptor = _serviceProvider?.GetService<RegistryOpenVerbInterceptor>();
                    interceptor?.DisableAndRestore();
                }
            }
            catch (Exception ex)
            {
                _logger?.Error("Failed to process Explorer open-verb interception state on exit.", ex);
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
            _singleInstanceService?.Dispose();
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

        string exePath = AppEnvironment.ResolveLaunchExecutablePath();
        var startupRegistrar = new StartupRegistrar("WinTab", exePath);
        services.AddSingleton(startupRegistrar);

        services.AddSingleton<IWindowManager, WindowManager>();
        services.AddSingleton<IWindowEventSource, WindowEventWatcher>();
        services.AddSingleton<ShellLocationIdentityService>();

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



    internal static void RequestExplicitShutdown()
    {
        if (_explicitShutdownRequested)
            return;

        _explicitShutdownRequested = true;
        Current.Shutdown();
    }

    public void SetTrayIconVisibility(bool visible)
    {
        SetTrayIconVisibilityCore(visible);
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




}
