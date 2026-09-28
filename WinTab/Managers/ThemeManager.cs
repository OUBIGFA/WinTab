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
        SetColor("TextPrimaryColor", dark ? "#F0F0F0" : "#202020");
        SetColor("TextSecondaryColor", dark ? "#BDBDBD" : "#595959");
        SetColor("TextTertiaryColor", dark ? "#909090" : "#707070");
        SetColor("TextAccentColor", dark ? "#F0F0F0" : "#181818");
        // Controls need stronger outlines than the quieter card borders and row dividers.
        SetColor("ControlBorderColor", dark ? "#8A8A8A" : "#888888");
        SetColor("BorderColor", dark ? "#454545" : "#C8C8C8");
        SetColor("DividerColor", dark ? "#3B3B3B" : "#D8D8D8");
        SetColor("ControlBackgroundColor", dark ? "#303030" : "#E9E9E9");
        SetColor("ControlHoverColor", dark ? "#3D3D3D" : "#DDDDDD");
        SetColor("DropdownBackgroundColor", dark ? "#1B1B1B" : "#FAFAFA");
        SetColor("ShadowColor", dark ? "#111111" : "#5A5A5A");
        SetColor("CheckBoxCheckedBackgroundColor", dark ? "#E5E5E5" : "#2A2A2A");
        SetColor("CheckBoxCheckedBorderColor", dark ? "#E5E5E5" : "#2A2A2A");
        SetColor("CheckBoxCheckedGlyphColor", dark ? "#242424" : "#EEEEEE");
        SetColor("FocusRingColor", dark ? "#C5C5C5" : "#444444");

        SetBrush("WindowBackgroundBrush", dark ? "#161616" : "#EDEDED");
        SetBrush("WindowTitleBarBrush", dark ? "#191919" : "#EDEDED");
        SetBrush("WindowBorderBrush", dark ? "#373737" : "#C7C7C7");
        SetBrush("SurfaceBrush", dark ? "#242424" : "#FCFCFC");
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
