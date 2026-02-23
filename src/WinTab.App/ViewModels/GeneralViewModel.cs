using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WinTab.Core.Enums;
using WinTab.Core.Interfaces;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.Platform.Win32;
using WinTab.UI.Localization;
using WinTab.App.Services;

namespace WinTab.App.ViewModels;

public partial class GeneralViewModel : ObservableObject
{
    private readonly AppSettings _settings;
    private readonly SettingsStore _settingsStore;
    private readonly StartupRegistrar _startupRegistrar;
    private readonly Logger _logger;
    private readonly RegistryOpenVerbInterceptor _openVerbInterceptor;

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

    [ObservableProperty]
    private bool _interceptExplorerFolderOpen;

    public GeneralViewModel(
        AppSettings settings,
        SettingsStore settingsStore,
        StartupRegistrar startupRegistrar,
        Logger logger,
        RegistryOpenVerbInterceptor openVerbInterceptor)
    {
        _settings = settings;
        _settingsStore = settingsStore;
        _startupRegistrar = startupRegistrar;
        _logger = logger;
        _openVerbInterceptor = openVerbInterceptor;

        // Load current values from settings
        _runAtStartup = startupRegistrar.IsEnabled();
        _startMinimized = settings.StartMinimized;
        _showTrayIcon = settings.EnableTrayIcon;
        _minimizeToTrayOnClose = settings.MinimizeToTrayOnClose;
        _selectedLanguageIndex = settings.Language == Language.Chinese ? 0 : 1;
        _restoreSession = settings.RestoreSessionOnStartup;
        _autoCloseEmpty = settings.AutoCloseEmptyGroups;
        _groupSameProcess = settings.GroupSameProcessWindows;
        _interceptExplorerFolderOpen = settings.EnableExplorerOpenVerbInterception;
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

    partial void OnInterceptExplorerFolderOpenChanged(bool value)
    {
        _settings.EnableExplorerOpenVerbInterception = value;
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
            _openVerbInterceptor.DisableAndRestore();
            _settings.EnableExplorerOpenVerbInterception = false;
            SaveSettings();

            _logger.Info("Explorer folder open behavior restored to system defaults.");
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to restore Explorer defaults.", ex);
        }
    }
}
