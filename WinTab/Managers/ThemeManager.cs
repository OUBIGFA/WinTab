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

        var dark = IsDarkTheme;

        SetColor("PrimaryColor", dark ? "#F2F2F2" : "#202020");
        SetColor("PrimaryLightColor", dark ? "#CFCFCF" : "#4A4A4A");
        SetColor("TextPrimaryColor", dark ? "#F2F2F2" : "#1E1E1E");
        SetColor("TextSecondaryColor", dark ? "#B8B8B8" : "#666666");
        SetColor("TextTertiaryColor", dark ? "#858585" : "#8A8A8A");
        SetColor("TextAccentColor", dark ? "#FFFFFF" : "#303030");
        SetColor("BorderColor", dark ? "#353535" : "#D7D7D7");
        SetColor("ControlBackgroundColor", dark ? "#202020" : "#F6F6F6");
        SetColor("ControlHoverColor", dark ? "#2A2A2A" : "#EDEDED");
        SetColor("DangerColor", dark ? "#B8B8B8" : "#555555");
        SetColor("DropdownBackgroundColor", dark ? "#181818" : "#FFFFFF");
        SetColor("ShadowColor", dark ? "#080808" : "#9A9A9A");
        SetColor("SuccessColor", dark ? "#B8B8B8" : "#4A4A4A");
        SetColor("WarningColor", dark ? "#B8B8B8" : "#4A4A4A");
        SetColor("AccentColor", dark ? "#F2F2F2" : "#202020");
        SetColor("StatusActiveColor", dark ? "#F2F2F2" : "#202020");
        SetColor("StatusInactiveColor", dark ? "#858585" : "#8A8A8A");
        SetColor("CheckBoxCheckedBackgroundColor", dark ? "#F2F2F2" : "#202020");
        SetColor("CheckBoxCheckedBorderColor", dark ? "#F2F2F2" : "#202020");
        SetColor("CheckBoxCheckedGlyphColor", dark ? "#181818" : "#FFFFFF");
        SetColor("StatusPillBackgroundColor", dark ? "#202020" : "#F1F1F1");
        SetColor("StatusPillBorderColor", dark ? "#353535" : "#D7D7D7");
        SetColor("FooterColor", dark ? "#161616" : "#F6F6F6");

        SetBrush("WindowBackgroundBrush", dark ? "#121212" : "#FAFAFA");
        SetBrush("WindowTitleBarBrush", dark ? "#161616" : "#FAFAFA");
        SetBrush("WindowBorderBrush", dark ? "#353535" : "#D7D7D7");
        SetBrush("SurfaceBrush", dark ? "#181818" : "#FFFFFF");
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
