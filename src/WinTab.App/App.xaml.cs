using System.Reflection;
using System.Threading;
using System.Windows;
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
using WinTab.TabHost;

namespace WinTab.App;

public partial class App : Application
{
    private static Mutex? _singleInstanceMutex;
    private IServiceProvider? _serviceProvider;
    private TrayIconController? _trayIconController;
    private Logger? _logger;
    private AppLifecycleService? _lifecycleService;
    private ExplorerTabHookService? _explorerTabHook;

    public static IServiceProvider Services { get; private set; } = null!;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // -- 1. Single instance check -------------------------------------
        _singleInstanceMutex = new Mutex(true, "WinTab_SingleInstance", out bool isNewInstance);
        if (!isNewInstance)
        {
            ActivateExistingInstance();
            Shutdown();
            return;
        }

        // -- 2. Install crash reporter ------------------------------------
        CrashReporter.Install(AppPaths.CrashLogPath);

        // -- 3. Create logger ---------------------------------------------
        _logger = new Logger(AppPaths.LogPath);
        _logger.Info("WinTab starting up...");

        // -- 4. Load settings ---------------------------------------------
        var settingsStore = new SettingsStore(AppPaths.SettingsPath, _logger);
        AppSettings settings = settingsStore.Load();

        // -- 5. Apply language --------------------------------------------
        LocalizationManager.ApplyLanguage(settings.Language);

        // -- 6. Apply theme -----------------------------------------------
        ThemeManager.ApplyTheme(settings.Theme);

        // -- 7. Configure DI ----------------------------------------------
        var services = new ServiceCollection();
        ConfigureServices(services, settings, settingsStore, _logger);
        _serviceProvider = services.BuildServiceProvider();
        Services = _serviceProvider;

        // -- 8. Create and show MainWindow --------------------------------
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

        // -- 9. Start lifecycle services ----------------------------------
        _lifecycleService = _serviceProvider.GetRequiredService<AppLifecycleService>();
        _lifecycleService.Start(settings);

        // Explorer integration (native Explorer tab conversion path).
        _explorerTabHook = _serviceProvider.GetRequiredService<ExplorerTabHookService>();

        // -- 10. Initialize tray icon -------------------------------------
        if (settings.EnableTrayIcon)
        {
            _trayIconController = _serviceProvider.GetRequiredService<TrayIconController>();
            _trayIconController.SetVisible(true);
        }

        _logger.Info("WinTab started successfully.");
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _logger?.Info("WinTab shutting down...");

        try
        {
            _lifecycleService?.Stop();
            _trayIconController?.Dispose();

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
            _singleInstanceMutex?.ReleaseMutex();
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

        var sessionStore = new SessionStore(AppPaths.SessionBackupPath, logger);
        services.AddSingleton(sessionStore);

        string exePath = Environment.ProcessPath ?? Assembly.GetExecutingAssembly().Location;
        var startupRegistrar = new StartupRegistrar("WinTab", exePath);
        services.AddSingleton(startupRegistrar);

        services.AddSingleton<IWindowManager, WindowManager>();
        services.AddSingleton<IWindowEventSource, WindowEventWatcher>();
        services.AddSingleton<IHotKeyManager, GlobalHotKeyManager>();
        services.AddSingleton<DragDetector>();

        services.AddSingleton<IGroupManager>(sp =>
            new TabGroupManager(
                sp.GetRequiredService<IWindowManager>(),
                sp.GetRequiredService<IWindowEventSource>(),
                sp.GetRequiredService<AppSettings>().TabBarHeight));

        services.AddSingleton<AppLifecycleService>();

        // Back to native Explorer-tab pipeline (not overlay hijack).
        services.AddSingleton<ExplorerTabHookService>();

        services.AddSingleton<TabHostBootstrapper>();

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
                        window.Activate();
                    }
                },
                exitApp: () => Current.Shutdown()));

        services.AddTransient<MainViewModel>();
        services.AddTransient<GeneralViewModel>();
        services.AddTransient<AppearanceViewModel>();
        services.AddTransient<AutoGroupingViewModel>();
        services.AddTransient<ShortcutsViewModel>();
        services.AddTransient<GroupsViewModel>();
        services.AddTransient<AboutViewModel>();

        services.AddSingleton<MainWindow>();
        services.AddTransient<GeneralPage>();
        services.AddTransient<AppearancePage>();
        services.AddTransient<AutoGroupingPage>();
        services.AddTransient<ShortcutsPage>();
        services.AddTransient<GroupsPage>();
        services.AddTransient<AboutPage>();
    }

    private static void ActivateExistingInstance()
    {
        try
        {
            using var current = System.Diagnostics.Process.GetCurrentProcess();
            var existing = System.Diagnostics.Process.GetProcessesByName(current.ProcessName)
                .FirstOrDefault(p => p.Id != current.Id);

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
}
