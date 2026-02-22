using System.Diagnostics;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WinTab.Core.Enums;
using WinTab.Core.Interfaces;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.Platform.Win32;
using WinTab.UI.Localization;
using Microsoft.Win32;

namespace WinTab.App.ViewModels;

public partial class GeneralViewModel : ObservableObject
{
    private readonly AppSettings _settings;
    private readonly SettingsStore _settingsStore;
    private readonly StartupRegistrar _startupRegistrar;
    private readonly Logger _logger;

    [ObservableProperty]
    private bool _runAtStartup;

    [ObservableProperty]
    private bool _startMinimized;

    [ObservableProperty]
    private bool _showTrayIcon;

    [ObservableProperty]
    private int _selectedLanguageIndex;

    [ObservableProperty]
    private bool _restoreSession;

    [ObservableProperty]
    private bool _autoCloseEmpty;

    [ObservableProperty]
    private bool _groupSameProcess;

    public GeneralViewModel(
        AppSettings settings,
        SettingsStore settingsStore,
        StartupRegistrar startupRegistrar,
        Logger logger)
    {
        _settings = settings;
        _settingsStore = settingsStore;
        _startupRegistrar = startupRegistrar;
        _logger = logger;

        // Load current values from settings
        _runAtStartup = startupRegistrar.IsEnabled();
        _startMinimized = settings.StartMinimized;
        _showTrayIcon = settings.EnableTrayIcon;
        _selectedLanguageIndex = settings.Language == Language.Chinese ? 0 : 1;
        _restoreSession = settings.RestoreSessionOnStartup;
        _autoCloseEmpty = settings.AutoCloseEmptyGroups;
        _groupSameProcess = settings.GroupSameProcessWindows;
    }

    partial void OnRunAtStartupChanged(bool value)
    {
        _startupRegistrar.SetEnabled(value);
        _settings.RunAtStartup = value;
        SaveSettings();
        _logger.Info($"RunAtStartup changed to {value}");
    }

    partial void OnStartMinimizedChanged(bool value)
    {
        _settings.StartMinimized = value;
        SaveSettings();
    }

    partial void OnShowTrayIconChanged(bool value)
    {
        _settings.EnableTrayIcon = value;
        SaveSettings();
    }

    partial void OnSelectedLanguageIndexChanged(int value)
    {
        var language = value == 0 ? Language.Chinese : Language.English;
        _settings.Language = language;
        LocalizationManager.ApplyLanguage(language);
        SaveSettings();
        _logger.Info($"Language changed to {language}");
    }

    partial void OnRestoreSessionChanged(bool value)
    {
        _settings.RestoreSessionOnStartup = value;
        SaveSettings();
    }

    partial void OnAutoCloseEmptyChanged(bool value)
    {
        _settings.AutoCloseEmptyGroups = value;
        SaveSettings();
    }

    partial void OnGroupSameProcessChanged(bool value)
    {
        _settings.GroupSameProcessWindows = value;
        SaveSettings();
    }

    private void SaveSettings()
    {
        _settingsStore.SaveDebounced(_settings);
    }

    [RelayCommand]
    private void RestoreExplorerDefaults()
    {
        try
        {
            string baseDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "WinTab");
            string backupDir = Path.Combine(baseDir, "reg-backups");
            Directory.CreateDirectory(backupDir);

            string ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string backupPath = Path.Combine(backupDir, $"HKCU_Software_Classes_Folder_shell_BEFORE_RESTORE_{ts}.reg");

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "reg.exe",
                    Arguments = $"export \"HKCU\\Software\\Classes\\Folder\\shell\" \"{backupPath}\" /y",
                    UseShellExecute = false,
                    CreateNoWindow = true
                })?.WaitForExit(2000);
            }
            catch
            {
                // ignore backup failures
            }

            using (RegistryKey? folderShell = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Folder\shell", writable: true))
            {
                // Force default verb back to Explorer's open action.
                folderShell?.SetValue(string.Empty, "open", RegistryValueKind.String);
            }

            // Remove per-user overrides that can hijack folder opening.
            using (RegistryKey? root = Registry.CurrentUser.OpenSubKey(@"Software\Classes", writable: true))
            {
                root?.DeleteSubKeyTree(@"Folder\shell\open\command", throwOnMissingSubKey: false);
                root?.DeleteSubKeyTree(@"Folder\shell\opennewtab\command", throwOnMissingSubKey: false);
                root?.DeleteSubKeyTree(@"Directory\shell\none", throwOnMissingSubKey: false);
                root?.DeleteSubKeyTree(@"Directory\shell\open", throwOnMissingSubKey: false);
                root?.DeleteSubKeyTree(@"Drive\shell\none", throwOnMissingSubKey: false);
                root?.DeleteSubKeyTree(@"Drive\shell\open", throwOnMissingSubKey: false);
            }

            // Restart Explorer so Shell picks up changes immediately.
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "taskkill.exe",
                    Arguments = "/f /im explorer.exe",
                    UseShellExecute = false,
                    CreateNoWindow = true
                })?.WaitForExit(4000);

                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    UseShellExecute = true
                });
            }
            catch
            {
                // ignore
            }

            _logger.Info("Explorer folder open behavior restored to default 'open' verb.");
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to restore Explorer defaults.", ex);
        }
    }
}
