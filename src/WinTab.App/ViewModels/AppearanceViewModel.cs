using CommunityToolkit.Mvvm.ComponentModel;
using WinTab.Core.Enums;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.UI.Themes;

namespace WinTab.App.ViewModels;

public partial class AppearanceViewModel : ObservableObject
{
    private readonly AppSettings _settings;
    private readonly SettingsStore _settingsStore;
    private readonly Logger _logger;
    private bool _isSynchronizingThemeSelection;

    // Theme radio button backing fields
    [ObservableProperty]
    private bool _isThemeLight;

    [ObservableProperty]
    private bool _isThemeDark;

    public AppearanceViewModel(
        AppSettings settings,
        SettingsStore settingsStore,
        Logger logger)
    {
        _settings = settings;
        _settingsStore = settingsStore;
        _logger = logger;

        // System theme mode is no longer user-selectable; treat as Light.
        if (settings.Theme == ThemeMode.System)
            settings.Theme = ThemeMode.Light;

        SynchronizeThemeSelection(settings.Theme);
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

        _settingsStore.SaveDebounced(_settings);

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
}
