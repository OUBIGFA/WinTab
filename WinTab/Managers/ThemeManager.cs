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

        SetColor("PrimaryColor", dark ? "#F0F0F0" : "#1F1F1F");
        SetColor("PrimaryLightColor", dark ? "#C9C9C9" : "#515151");
        SetColor("TextPrimaryColor", dark ? "#F0F0F0" : "#1D1D1D");
        SetColor("TextSecondaryColor", dark ? "#BDBDBD" : "#565656");
        SetColor("TextTertiaryColor", dark ? "#9A9A9A" : "#767676");
        SetColor("TextAccentColor", dark ? "#FFFFFF" : "#111111");
        SetColor("BorderColor", dark ? "#343434" : "#D4D4D4");
        SetColor("DividerColor", dark ? "#2A2A2A" : "#E1E1E1");
        SetColor("ControlBackgroundColor", dark ? "#202020" : "#F2F2F2");
        SetColor("ControlHoverColor", dark ? "#2A2A2A" : "#E9E9E9");
        SetColor("DangerColor", dark ? "#D0D0D0" : "#3A3A3A");
        SetColor("DropdownBackgroundColor", dark ? "#181818" : "#FDFDFD");
        SetColor("ShadowColor", dark ? "#050505" : "#6A6A6A");
        SetColor("SuccessColor", dark ? "#D0D0D0" : "#3A3A3A");
        SetColor("WarningColor", dark ? "#D0D0D0" : "#3A3A3A");
        SetColor("AccentColor", dark ? "#F0F0F0" : "#1F1F1F");
        SetColor("StatusActiveColor", dark ? "#F0F0F0" : "#1F1F1F");
        SetColor("StatusInactiveColor", dark ? "#9A9A9A" : "#767676");
        SetColor("CheckBoxCheckedBackgroundColor", dark ? "#F0F0F0" : "#1F1F1F");
        SetColor("CheckBoxCheckedBorderColor", dark ? "#F0F0F0" : "#1F1F1F");
        SetColor("CheckBoxCheckedGlyphColor", dark ? "#171717" : "#F8F8F8");
        SetColor("StatusPillBackgroundColor", dark ? "#232323" : "#EEEEEE");
        SetColor("StatusPillBorderColor", dark ? "#393939" : "#D4D4D4");
        SetColor("FocusRingColor", dark ? "#D0D0D0" : "#3A3A3A");
        SetColor("FooterColor", dark ? "#181818" : "#F2F2F2");

        SetBrush("WindowBackgroundBrush", dark ? "#121212" : "#F7F7F7");
        SetBrush("WindowTitleBarBrush", dark ? "#161616" : "#F7F7F7");
        SetBrush("WindowBorderBrush", dark ? "#343434" : "#D4D4D4");
        SetBrush("SurfaceBrush", dark ? "#181818" : "#FDFDFD");
        SetBrush("SurfaceMutedBrush", dark ? "#202020" : "#F2F2F2");
        SetBrush("SurfaceRaisedBrush", dark ? "#1C1C1C" : "#FFFFFF");
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
