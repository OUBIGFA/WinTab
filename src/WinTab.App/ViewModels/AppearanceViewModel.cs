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

    // Theme radio button backing fields
    [ObservableProperty]
    private bool _isThemeSystem;

    [ObservableProperty]
    private bool _isThemeLight;

    [ObservableProperty]
    private bool _isThemeDark;

    // Tab style radio button backing fields
    [ObservableProperty]
    private bool _isTabStyleModern;

    [ObservableProperty]
    private bool _isTabStyleTraditional;

    [ObservableProperty]
    private bool _isTabStyleCompact;

    // Tab bar settings
    [ObservableProperty]
    private bool _useRoundedCorners;

    [ObservableProperty]
    private bool _useMicaEffect;

    [ObservableProperty]
    private int _tabBarHeight;

    [ObservableProperty]
    private double _tabBarOpacity;

    public AppearanceViewModel(
        AppSettings settings,
        SettingsStore settingsStore,
        Logger logger)
    {
        _settings = settings;
        _settingsStore = settingsStore;
        _logger = logger;

        // Initialize theme radio buttons
        _isThemeSystem = settings.Theme == ThemeMode.System;
        _isThemeLight = settings.Theme == ThemeMode.Light;
        _isThemeDark = settings.Theme == ThemeMode.Dark;

        // Initialize tab style radio buttons
        _isTabStyleModern = settings.TabStyle == TabStyle.Modern;
        _isTabStyleTraditional = settings.TabStyle == TabStyle.Traditional;
        _isTabStyleCompact = settings.TabStyle == TabStyle.Compact;

        // Initialize tab bar settings
        _useRoundedCorners = settings.UseRoundedCorners;
        _useMicaEffect = settings.UseMicaEffect;
        _tabBarHeight = settings.TabBarHeight;
        _tabBarOpacity = settings.TabBarOpacity;
    }

    partial void OnIsThemeSystemChanged(bool value)
    {
        if (value) ApplyTheme(ThemeMode.System);
    }

    partial void OnIsThemeLightChanged(bool value)
    {
        if (value) ApplyTheme(ThemeMode.Light);
    }

    partial void OnIsThemeDarkChanged(bool value)
    {
        if (value) ApplyTheme(ThemeMode.Dark);
    }

    partial void OnIsTabStyleModernChanged(bool value)
    {
        if (value) ApplyTabStyle(TabStyle.Modern);
    }

    partial void OnIsTabStyleTraditionalChanged(bool value)
    {
        if (value) ApplyTabStyle(TabStyle.Traditional);
    }

    partial void OnIsTabStyleCompactChanged(bool value)
    {
        if (value) ApplyTabStyle(TabStyle.Compact);
    }

    partial void OnUseRoundedCornersChanged(bool value)
    {
        _settings.UseRoundedCorners = value;
        SaveSettings();
    }

    partial void OnUseMicaEffectChanged(bool value)
    {
        _settings.UseMicaEffect = value;
        SaveSettings();
    }

    partial void OnTabBarHeightChanged(int value)
    {
        _settings.TabBarHeight = value;
        SaveSettings();
    }

    partial void OnTabBarOpacityChanged(double value)
    {
        _settings.TabBarOpacity = value;
        SaveSettings();
    }

    private void ApplyTheme(ThemeMode mode)
    {
        _settings.Theme = mode;
        ThemeManager.ApplyTheme(mode);
        SaveSettings();
        _logger.Info($"Theme changed to {mode}");
    }

    private void ApplyTabStyle(TabStyle style)
    {
        _settings.TabStyle = style;
        SaveSettings();
        _logger.Info($"Tab style changed to {style}");
    }

    private void SaveSettings()
    {
        _settingsStore.SaveDebounced(_settings);
    }
}
