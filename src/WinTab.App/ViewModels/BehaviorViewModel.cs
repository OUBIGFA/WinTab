using CommunityToolkit.Mvvm.ComponentModel;
using WinTab.App.ExplorerTabUtilityPort;
using WinTab.Core.Models;
using WinTab.Persistence;

namespace WinTab.App.ViewModels;

public partial class BehaviorViewModel : ObservableObject
{
    private readonly AppSettings _settings;
    private readonly SettingsStore _settingsStore;
    private readonly ExplorerTabMouseHookService _tabMouseHookService;
    private readonly IExplorerAutoConvertController _autoConvertController;
    private bool _isUpdatingCloseTabOnDoubleClickToggle;

    [ObservableProperty]
    private bool _openNewTabFromActiveTabPath;

    [ObservableProperty]
    private bool _openChildFolderInNewTabFromActiveTab;

    [ObservableProperty]
    private bool _closeTabOnDoubleClick;

    [ObservableProperty]
    private bool _enableAutoConvertExplorerWindows;

    public bool IsOpenChildFolderInNewTabOptionEnabled => EnableAutoConvertExplorerWindows;

    public BehaviorViewModel(
        AppSettings settings,
        SettingsStore settingsStore,
        ExplorerTabMouseHookService tabMouseHookService,
        IExplorerAutoConvertController autoConvertController)
    {
        _settings = settings;
        _settingsStore = settingsStore;
        _tabMouseHookService = tabMouseHookService;
        _autoConvertController = autoConvertController;

        _openNewTabFromActiveTabPath = settings.OpenNewTabFromActiveTabPath;
        _openChildFolderInNewTabFromActiveTab = settings.OpenChildFolderInNewTabFromActiveTab;
        _closeTabOnDoubleClick = settings.CloseTabOnDoubleClick;
        _enableAutoConvertExplorerWindows = settings.EnableAutoConvertExplorerWindows;
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

    partial void OnEnableAutoConvertExplorerWindowsChanged(bool value)
    {
        _settings.EnableAutoConvertExplorerWindows = value;
        _autoConvertController.SetAutoConvertEnabled(value);
        SaveSettings();
        OnPropertyChanged(nameof(IsOpenChildFolderInNewTabOptionEnabled));
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
