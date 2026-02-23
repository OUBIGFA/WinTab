using System.Reflection;
using System.IO;
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
    private const int DefaultTabBarHeight = 32;

    private static Mutex? _singleInstanceMutex;
    private static bool _ownsSingleInstanceMutex;
    private static bool _explicitShutdownRequested;
    private IServiceProvider? _serviceProvider;
    private TrayIconController? _trayIconController;
    private Logger? _logger;
    private AppLifecycleService? _lifecycleService;
    private ExplorerTabHookService? _explorerTabHook;
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
            ActivateExistingInstance();
            Shutdown();
            return;
        }

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

        // -- 10. Start lifecycle services ---------------------------------
        _lifecycleService = _serviceProvider.GetRequiredService<AppLifecycleService>();
        _lifecycleService.Start(settings);

        // Explorer integration (native Explorer tab conversion path).
        _explorerTabHook = _serviceProvider.GetRequiredService<ExplorerTabHookService>();

        // IPC handler: allow handler invocations to forward open-folder requests.
        _openRequestServer = _serviceProvider.GetRequiredService<ExplorerOpenRequestServer>();
        _openRequestServer.Start(async path =>
        {
            try
            {
                var hook = _serviceProvider.GetRequiredService<ExplorerTabHookService>();
                await hook.OpenLocationAsTabAsync(path);
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
            bool enableExplorerOpenVerbInterception = isWin11;

            // This behavior is now product default on supported OS.
            settings.EnableExplorerOpenVerbInterception = enableExplorerOpenVerbInterception;

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
        SetTrayIconVisibilityCore(settings.EnableTrayIcon);

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

        var sessionStore = new SessionStore(AppPaths.SessionBackupPath, logger);
        services.AddSingleton(sessionStore);

        string exePath = ResolveLaunchExecutablePath();
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
                DefaultTabBarHeight));

        services.AddSingleton<AppLifecycleService>();

        services.AddSingleton(sp =>
            new RegistryOpenVerbInterceptor(
                exePath,
                sp.GetRequiredService<Logger>()));

        services.AddSingleton<ExplorerOpenRequestServer>();

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
                exitApp: RequestExplicitShutdown));

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

    private static bool TryHandleOpenFolderInvocation(string[] args, Logger? logger)
    {
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
            logger?.Info($"Forwarded open-folder request to existing instance: {path}");
            return true;
        }

        logger?.Warn($"No existing instance pipe; falling back to Explorer open: {path}");
        TryOpenFolderFallback(path, logger);

        return true;
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

    internal static void SetTrayIconVisibility(bool visible)
    {
        if (Current is App app)
            app.SetTrayIconVisibilityCore(visible);
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
}
