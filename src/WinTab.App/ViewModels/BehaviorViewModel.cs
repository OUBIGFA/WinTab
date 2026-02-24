using CommunityToolkit.Mvvm.ComponentModel;
using WinTab.App.ExplorerTabUtilityPort;
using WinTab.App.Services;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Persistence;

namespace WinTab.App.ViewModels;

public partial class BehaviorViewModel : ObservableObject
{
    private readonly AppSettings _settings;
    private readonly SettingsStore _settingsStore;
    private readonly RegistryOpenVerbInterceptor _openVerbInterceptor;
    private readonly ExplorerTabMouseHookService _tabMouseHookService;
    private readonly Logger _logger;
    private bool _isUpdatingExplorerOpenVerbToggle;
    private bool _isUpdatingCloseTabOnDoubleClickToggle;

    [ObservableProperty]
    private bool _openNewTabFromActiveTabPath;

    [ObservableProperty]
    private bool _openChildFolderInNewTabFromActiveTab;

    [ObservableProperty]
    private bool _enableExplorerOpenVerbInterception;

    [ObservableProperty]
    private bool _closeTabOnDoubleClick;

    public bool IsOpenChildFolderInNewTabOptionEnabled => EnableExplorerOpenVerbInterception;

    public BehaviorViewModel(
        AppSettings settings,
        SettingsStore settingsStore,
        RegistryOpenVerbInterceptor openVerbInterceptor,
        ExplorerTabMouseHookService tabMouseHookService,
        Logger logger)
    {
        _settings = settings;
        _settingsStore = settingsStore;
        _openVerbInterceptor = openVerbInterceptor;
        _tabMouseHookService = tabMouseHookService;
        _logger = logger;

        _openNewTabFromActiveTabPath = settings.OpenNewTabFromActiveTabPath;
        _openChildFolderInNewTabFromActiveTab = settings.OpenChildFolderInNewTabFromActiveTab;
        _enableExplorerOpenVerbInterception = settings.EnableExplorerOpenVerbInterception;
        _closeTabOnDoubleClick = settings.CloseTabOnDoubleClick;
    }

    partial void OnOpenNewTabFromActiveTabPathChanged(bool value)
    {
        _settings.OpenNewTabFromActiveTabPath = value;
        SaveSettings();
    }

    partial void OnOpenChildFolderInNewTabFromActiveTabChanged(bool value)
    {
        _settings.OpenChildFolderInNewTabFromActiveTab = value;
        SaveSettings();
    }

    partial void OnEnableExplorerOpenVerbInterceptionChanged(bool value)
    {
        if (_isUpdatingExplorerOpenVerbToggle)
            return;

        if (value && !OperatingSystem.IsWindowsVersionAtLeast(10, 0, 22000))
        {
            _logger.Warn("Explorer open-verb interception requires Windows 11; toggle ignored.");
            SetExplorerOpenVerbToggle(false);
            return;
        }

        try
        {
            if (value)
                _openVerbInterceptor.EnableOrRepair();
            else
                _openVerbInterceptor.DisableAndRestore();

            _settings.EnableExplorerOpenVerbInterception = value;
            SaveSettings();

            OnPropertyChanged(nameof(IsOpenChildFolderInNewTabOptionEnabled));

            _logger.Info($"Explorer open-verb interception toggled to {(value ? "enabled" : "disabled")}.");
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to apply Explorer open-verb interception toggle.", ex);
            SetExplorerOpenVerbToggle(_settings.EnableExplorerOpenVerbInterception);
        }
    }

    partial void OnCloseTabOnDoubleClickChanged(bool value)
    {
        if (_isUpdatingCloseTabOnDoubleClickToggle)
            return;

        bool applied = _tabMouseHookService.SetEnabled(value);
        if (applied != value)
        {
            SetCloseTabOnDoubleClickToggle(applied);
            return;
        }

        _settings.CloseTabOnDoubleClick = applied;
        SaveSettings();
    }

    private void SetExplorerOpenVerbToggle(bool value)
    {
        _isUpdatingExplorerOpenVerbToggle = true;
        try
        {
            EnableExplorerOpenVerbInterception = value;
            _settings.EnableExplorerOpenVerbInterception = value;
            SaveSettings();

            OnPropertyChanged(nameof(IsOpenChildFolderInNewTabOptionEnabled));
        }
        finally
        {
            _isUpdatingExplorerOpenVerbToggle = false;
        }
    }

    private void SetCloseTabOnDoubleClickToggle(bool value)
    {
        _isUpdatingCloseTabOnDoubleClickToggle = true;
        try
        {
            CloseTabOnDoubleClick = value;
            _settings.CloseTabOnDoubleClick = value;
            SaveSettings();
        }
        finally
        {
            _isUpdatingCloseTabOnDoubleClickToggle = false;
        }
    }

    private void SaveSettings()
    {
        _settingsStore.SaveDebounced(_settings);
    }
}
