using System;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace WinTab.Managers;

public static class ThemeManager
{
    private const string PersonalizeKey = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
    private static bool _followingSystem;

    /// <summary>Raised on the UI thread after the palette changes.</summary>
    public static event Action? ThemeApplied;

    public static bool IsDarkTheme => SettingsManager.Theme switch
    {
        "Dark" => true,
        "System" => SystemUsesDarkTheme(),
        _ => false
    };

    public static void ApplyTheme()
    {
        if (Application.Current == null)
            return;

        FollowSystem(string.Equals(SettingsManager.Theme, "System", StringComparison.Ordinal));
        ApplyTheme(IsDarkTheme);
        ThemeApplied?.Invoke();
    }

    /// <summary>Windows stores the app mode as AppsUseLightTheme; 0 selects dark and a missing value means light.</summary>
    internal static bool SystemUsesDarkTheme()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(PersonalizeKey);
            return key?.GetValue("AppsUseLightTheme") is int light && light == 0;
        }
        catch (Exception error) when (error is System.Security.SecurityException or UnauthorizedAccessException or System.IO.IOException)
        {
            return false;
        }
    }

    // Listen only while the preference follows the system, so fixed themes keep no extra subscription.
    private static void FollowSystem(bool follow)
    {
        if (follow == _followingSystem)
            return;
        _followingSystem = follow;
        if (follow) SystemEvents.UserPreferenceChanged += OnUserPreferenceChanged;
        else SystemEvents.UserPreferenceChanged -= OnUserPreferenceChanged;
    }

    private static void OnUserPreferenceChanged(object? sender, UserPreferenceChangedEventArgs args)
    {
        // Switching the app mode raises the General category for ImmersiveColorSet.
        if (args.Category != UserPreferenceCategory.General)
            return;
        Application.Current?.Dispatcher.InvokeAsync(() =>
        {
            if (_followingSystem) ApplyTheme();
        });
    }

    internal static void ApplyTheme(bool dark)
    {
        if (Application.Current == null)
            return;

        SetColor("PrimaryColor", dark ? "#E5E5E5" : "#171717");
        SetColor("PrimaryLightColor", dark ? "#BDBDBD" : "#404040");
        SetColor("TextPrimaryColor", dark ? "#FAFAFA" : "#0A0A0A");
        SetColor("TextSecondaryColor", dark ? "#A3A3A3" : "#707070");
        SetColor("TextTertiaryColor", dark ? "#909090" : "#737373");
        SetColor("TextAccentColor", dark ? "#FAFAFA" : "#171717");
        // Controls need stronger outlines than the quieter card borders and row dividers.
        SetColor("ControlBorderColor", dark ? "#858585" : "#888888");
        SetColor("BorderColor", dark ? "#333333" : "#E5E5E5");
        SetColor("DividerColor", dark ? "#333333" : "#E5E5E5");
        SetColor("ControlBackgroundColor", dark ? "#262626" : "#F5F5F5");
        SetColor("ControlHoverColor", dark ? "#333333" : "#EAEAEA");
        SetColor("DropdownBackgroundColor", dark ? "#171717" : "#FFFFFF");
        SetColor("ShadowColor", dark ? "#111111" : "#5A5A5A");
        SetColor("CheckBoxCheckedBackgroundColor", dark ? "#E5E5E5" : "#171717");
        SetColor("CheckBoxCheckedBorderColor", dark ? "#E5E5E5" : "#171717");
        SetColor("CheckBoxCheckedGlyphColor", dark ? "#171717" : "#FAFAFA");
        SetColor("FocusRingColor", dark ? "#A3A3A3" : "#737373");

        SetColor("SwitchOffColor", dark ? "#737373" : "#888888");

        SetColor("SwitchThumbColor", dark ? "#FAFAFA" : "#FFFFFF");

        SetBrush("WindowBackgroundBrush", dark ? "#0A0A0A" : "#FFFFFF");
        SetBrush("WindowTitleBarBrush", dark ? "#171717" : "#FAFAFA");
        SetBrush("WindowBorderBrush", dark ? "#333333" : "#E5E5E5");
        SetBrush("SurfaceBrush", dark ? "#171717" : "#FFFFFF");
    }

    private static void SetColor(string key, string hex)
    {
        var updated = (Color)ColorConverter.ConvertFromString(hex);
        Application.Current.Resources[key] = updated;

        var brushKey = key.Replace("Color", "Brush", StringComparison.Ordinal);
        Application.Current.Resources[brushKey] = new SolidColorBrush(updated);
    }

    private static void SetBrush(string key, string hex)
    {
        var updated = (Color)ColorConverter.ConvertFromString(hex);
        Application.Current.Resources[key] = new SolidColorBrush(updated);
    }
}
