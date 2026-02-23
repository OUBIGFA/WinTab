using System.Windows;
using System.Windows.Media;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;

using AppThemeMode = WinTab.Core.Enums.ThemeMode;

namespace WinTab.UI.Themes;

public static class ThemeManager
{
    private static AppThemeMode _currentMode = AppThemeMode.Light;

    private static readonly object ThemeApplyLock = new();

    public static AppThemeMode CurrentMode => _currentMode;

    public static bool IsDarkMode => _currentMode switch
    {
        AppThemeMode.Dark => true,
        AppThemeMode.Light => false,
        _ => ApplicationThemeManager.GetAppTheme() == ApplicationTheme.Dark
    };

    public static void ApplyTheme(AppThemeMode mode)
    {
        lock (ThemeApplyLock)
        {
            _currentMode = mode;

            // Only explicit Light/Dark is supported.
            var theme = mode == AppThemeMode.Dark
                ? ApplicationTheme.Dark
                : ApplicationTheme.Light;

            RunOnUiThread(() =>
            {
                ApplicationThemeManager.Apply(theme, WindowBackdropType.None, updateAccent: true);
            });
        }
    }

    private static void RunOnUiThread(Action action)
    {
        Application? app = Application.Current;
        if (app?.Dispatcher is null)
        {
            action();
            return;
        }

        if (app.Dispatcher.CheckAccess())
        {
            action();
            return;
        }

        app.Dispatcher.Invoke(action);
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
