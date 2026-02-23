using CommunityToolkit.Mvvm.ComponentModel;
using WinTab.Core.Enums;
using WinTab.Core.Interfaces;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.Platform.Win32;
using WinTab.UI.Localization;

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
    private bool _minimizeToTrayOnClose;

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
        _minimizeToTrayOnClose = settings.MinimizeToTrayOnClose;
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
        App.SetTrayIconVisibility(value);

        if (!value && MinimizeToTrayOnClose)
        {
            MinimizeToTrayOnClose = false;
            return;
        }

        if (value && !MinimizeToTrayOnClose)
        {
            MinimizeToTrayOnClose = true;
            return;
        }

        SaveSettings();
    }

    partial void OnMinimizeToTrayOnCloseChanged(bool value)
    {
        _settings.MinimizeToTrayOnClose = value;
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
}
