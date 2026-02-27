using CommunityToolkit.Mvvm.ComponentModel;
using WinTab.Core.Enums;
using WinTab.Core.Interfaces;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.Platform.Win32;
using WinTab.UI.Localization;
using WinTab.UI.Themes;
using WinTab.App.Services;

namespace WinTab.App.ViewModels;

public partial class GeneralViewModel : ObservableObject
{
    private readonly AppSettings _settings;
    private readonly SettingsStore _settingsStore;
    private readonly StartupRegistrar _startupRegistrar;
    private readonly Logger _logger;
    private readonly TrayIconController _trayIconController;
    private bool _isSynchronizingThemeSelection;

    [ObservableProperty]
    private bool _runAtStartup;

    [ObservableProperty]
    private bool _startMinimized;

    [ObservableProperty]
    private bool _showTrayIcon;

    [ObservableProperty]
    private int _selectedLanguageIndex;

    [ObservableProperty]
    private bool _isThemeLight;

    [ObservableProperty]
    private bool _isThemeDark;

    public GeneralViewModel(
        AppSettings settings,
        SettingsStore settingsStore,
        StartupRegistrar startupRegistrar,
        Logger logger,
        TrayIconController trayIconController)
    {
        _settings = settings;
        _settingsStore = settingsStore;
        _startupRegistrar = startupRegistrar;
        _logger = logger;
        _trayIconController = trayIconController;

        // Load current values from settings
        _runAtStartup = startupRegistrar.IsEnabled();
        _startMinimized = settings.StartMinimized;
        _showTrayIcon = settings.ShowTrayIcon;
        _selectedLanguageIndex = settings.Language == Language.Chinese ? 0 : 1;

        if (settings.Theme == ThemeMode.System)
            settings.Theme = ThemeMode.Light;

        SynchronizeThemeSelection(settings.Theme);
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
        _settings.ShowTrayIcon = value;
        SaveSettings();
        
        _trayIconController.SetVisible(value);
    }

    partial void OnSelectedLanguageIndexChanged(int value)
    {
        var language = value == 0 ? Language.Chinese : Language.English;
        _settings.Language = language;
        LocalizationManager.ApplyLanguage(language);
        SaveSettings();
        _logger.Info($"Language changed to {language}");
    }

    partial void OnIsThemeLightChanged(bool value)
    {
        if (!value || _isSynchronizingThemeSelection)
            return;

        ApplyTheme(ThemeMode.Light);
    }

    partial void OnIsThemeDarkChanged(bool value)
    {
        if (!value || _isSynchronizingThemeSelection)
            return;

        ApplyTheme(ThemeMode.Dark);
    }

    private void ApplyTheme(ThemeMode mode)
    {
        if (_settings.Theme == mode)
            return;

        _settings.Theme = mode;
        ThemeManager.ApplyTheme(mode);
        SynchronizeThemeSelection(mode);

        SaveSettings();
        _logger.Info($"Theme changed to {mode}");
    }

    private void SynchronizeThemeSelection(ThemeMode mode)
    {
        _isSynchronizingThemeSelection = true;
        try
        {
            IsThemeLight = mode == ThemeMode.Light;
            IsThemeDark = mode == ThemeMode.Dark;
        }
        finally
        {
            _isSynchronizingThemeSelection = false;
        }
    }

    private void SaveSettings()
    {
        _settingsStore.SaveDebounced(_settings);
    }
}
