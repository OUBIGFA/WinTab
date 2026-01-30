using System.Diagnostics;
using System.Windows;
using WinTab.App.Localization;
using WinTab.Core;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.Platform.Win32;

namespace WinTab.App;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : System.Windows.Application
{
    private const string AppName = "WinTab";

    private TrayIconController? _trayIcon;
    private Logger? _logger;
    private SettingsStore? _settingsStore;
    private StartupRegistrar? _startupRegistrar;
    private AppSettings _settings = new();
    private MainWindow? _mainWindow;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        _logger = new Logger(AppPaths.LogPath);
        _settingsStore = new SettingsStore(AppPaths.SettingsPath);
        _settings = _settingsStore.Load();
        AppSettings.CurrentInstance = _settings;

        var executablePath = GetExecutablePath();
        _startupRegistrar = new StartupRegistrar(AppName, executablePath);

        try
        {
            _startupRegistrar.SetEnabled(_settings.RunAtStartup);
        }
        catch (Exception ex)
        {
            _logger.Warn($"Startup registration failed: {ex.Message}");
        }

        // Apply language settings
        LocalizationManager.ApplyLanguage(_settings.Language);
        
        // Check for crash recovery backup on startup
        var hasBackup = _settingsStore.LoadWindowAttachments().Count > 0;
        if (hasBackup)
        {
            var result = System.Windows.MessageBox.Show(
                LocalizationManager.GetString("SessionRecovery_Message"),
                LocalizationManager.GetString("SessionRecovery_Title"),
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
                
            if (result == MessageBoxResult.No)
            {
                _settingsStore.ClearBackup();
                _settings.WindowAttachments.Clear();
            }
        }
        
        _mainWindow = new MainWindow(_settings, _settingsStore, _startupRegistrar, _logger, SetTrayVisible);

        _trayIcon = new TrayIconController(ShowMainWindow, ExitApplication);
        _trayIcon.SetVisible(_settings.EnableTrayIcon);

        if (!_settings.StartMinimized || !_settings.EnableTrayIcon)
        {
            _mainWindow.Show();
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        // Backup session before exiting for crash recovery
        _settingsStore?.BackupSession();
        
        _mainWindow?.Dispose();
        _trayIcon?.Dispose();
        _logger?.Dispose();
        base.OnExit(e);
    }

    private void ShowMainWindow()
    {
        if (_mainWindow is null)
        {
            return;
        }

        if (!_mainWindow.IsVisible)
        {
            _mainWindow.Show();
        }

        _mainWindow.Activate();
        _mainWindow.WindowState = WindowState.Normal;
    }

    private void ExitApplication()
    {
        _logger?.Info("Exit requested");
        Shutdown();
    }

    private void SetTrayVisible(bool visible)
    {
        _trayIcon?.SetVisible(visible);
    }

    private static string GetExecutablePath()
    {
        var modulePath = Process.GetCurrentProcess().MainModule?.FileName;
        if (!string.IsNullOrWhiteSpace(modulePath))
        {
            return modulePath;
        }

        return Environment.ProcessPath ?? string.Empty;
    }
}
