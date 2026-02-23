using CommunityToolkit.Mvvm.ComponentModel;
using WinTab.Core.Models;
using WinTab.Persistence;

namespace WinTab.App.ViewModels;

public partial class BehaviorViewModel : ObservableObject
{
    private readonly AppSettings _settings;
    private readonly SettingsStore _settingsStore;

    [ObservableProperty]
    private bool _restoreSession;

    [ObservableProperty]
    private bool _autoCloseEmpty;

    [ObservableProperty]
    private bool _groupSameProcess;

    public BehaviorViewModel(AppSettings settings, SettingsStore settingsStore)
    {
        _settings = settings;
        _settingsStore = settingsStore;

        _restoreSession = settings.RestoreSessionOnStartup;
        _autoCloseEmpty = settings.AutoCloseEmptyGroups;
        _groupSameProcess = settings.GroupSameProcessWindows;
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
