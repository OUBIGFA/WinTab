using System;
using System.Windows;
using System.Windows.Media;

namespace WinTab.Managers;

public static class ThemeManager
{
    public static bool IsDarkTheme => string.Equals(SettingsManager.Theme, "Dark", StringComparison.OrdinalIgnoreCase);

    public static void ApplyTheme()
    {
        if (Application.Current == null)
            return;

        ApplyTheme(IsDarkTheme);
    }

    internal static void ApplyTheme(bool dark)
    {
        if (Application.Current == null)
            return;

        SetColor("PrimaryColor", dark ? "#E8E8E8" : "#2A2A2A");
        SetColor("PrimaryLightColor", dark ? "#BDBDBD" : "#5E5E5E");
        SetColor("TextPrimaryColor", dark ? "#E5E5E5" : "#252525");
        SetColor("TextSecondaryColor", dark ? "#B2B2B2" : "#606060");
        SetColor("TextTertiaryColor", dark ? "#909090" : "#7A7A7A");
        SetColor("TextAccentColor", dark ? "#F0F0F0" : "#181818");
        SetColor("BorderColor", dark ? "#383838" : "#CCCCCC");
        SetColor("DividerColor", dark ? "#2B2B2B" : "#DEDEDE");
        SetColor("ControlBackgroundColor", dark ? "#242424" : "#EFEFEF");
        SetColor("ControlHoverColor", dark ? "#303030" : "#E4E4E4");
        SetColor("DropdownBackgroundColor", dark ? "#1B1B1B" : "#FAFAFA");
        SetColor("ShadowColor", dark ? "#111111" : "#5A5A5A");
        SetColor("CheckBoxCheckedBackgroundColor", dark ? "#E5E5E5" : "#2A2A2A");
        SetColor("CheckBoxCheckedBorderColor", dark ? "#E5E5E5" : "#2A2A2A");
        SetColor("CheckBoxCheckedGlyphColor", dark ? "#242424" : "#EEEEEE");
        SetColor("StatusPillBackgroundColor", dark ? "#303030" : "#E8E8E8");
        SetColor("StatusPillBorderColor", dark ? "#444444" : "#CCCCCC");
        SetColor("FocusRingColor", dark ? "#C5C5C5" : "#444444");

        SetBrush("WindowBackgroundBrush", dark ? "#151515" : "#F3F3F3");
        SetBrush("WindowTitleBarBrush", dark ? "#191919" : "#F3F3F3");
        SetBrush("WindowBorderBrush", dark ? "#373737" : "#C7C7C7");
        SetBrush("SurfaceBrush", dark ? "#1D1D1D" : "#F8F8F8");
        SetBrush("SurfaceMutedBrush", dark ? "#252525" : "#E9E9E9");
        SetBrush("SurfaceRaisedBrush", dark ? "#202020" : "#F6F6F6");
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
