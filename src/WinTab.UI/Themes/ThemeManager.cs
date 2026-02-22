using System.Windows;
using System.Windows.Media;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;

using AppThemeMode = WinTab.Core.Enums.ThemeMode;

namespace WinTab.UI.Themes;

public static class ThemeManager
{
    private static AppThemeMode _currentMode = AppThemeMode.System;

    public static AppThemeMode CurrentMode => _currentMode;

    public static bool IsDarkMode => _currentMode switch
    {
        AppThemeMode.Dark => true,
        AppThemeMode.Light => false,
        _ => ApplicationThemeManager.GetAppTheme() == ApplicationTheme.Dark
    };

    public static void ApplyTheme(AppThemeMode mode)
    {
        _currentMode = mode;

        var theme = mode switch
        {
            AppThemeMode.Light => ApplicationTheme.Light,
            AppThemeMode.Dark => ApplicationTheme.Dark,
            _ => ApplicationTheme.Unknown // System default
        };

        if (theme == ApplicationTheme.Unknown)
        {
            ApplicationThemeManager.ApplySystemTheme();
        }
        else
        {
            ApplicationThemeManager.Apply(theme);
        }
    }

    public static void ApplyMica(Window window, bool enabled)
    {
        if (!enabled) return;

        try
        {
            WindowBackdropType backdropType = WindowBackdropType.Mica;
            WindowBackdrop.ApplyBackdrop(window, backdropType);
        }
        catch
        {
            // Mica not supported on this OS version - ignore
        }
    }

    public static Color GetAccentColor()
    {
        try
        {
            return SystemParameters.WindowGlassColor;
        }
        catch
        {
            return Colors.CornflowerBlue;
        }
    }
}
