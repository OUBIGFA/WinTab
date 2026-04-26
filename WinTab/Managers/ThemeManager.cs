using System;
using System.Linq;
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

        SetColor("PrimaryColor", dark ? "#D6A74B" : "#303741");
        SetColor("PrimaryLightColor", dark ? "#E9C879" : "#596575");
        SetColor("TextPrimaryColor", dark ? "#EEF2F6" : "#20262D");
        SetColor("TextSecondaryColor", dark ? "#B7C0CC" : "#66717E");
        SetColor("TextTertiaryColor", dark ? "#8B97A5" : "#87909B");
        SetColor("TextAccentColor", dark ? "#F0D08A" : "#2F3B48");
        SetColor("BorderColor", dark ? "#3F4853" : "#D5DAE1");
        SetColor("ControlBackgroundColor", dark ? "#202832" : "#F8FAFC");
        SetColor("ControlHoverColor", dark ? "#2B3541" : "#EEF2F6");
        SetColor("DangerColor", dark ? "#F06A6A" : "#C93F3F");
        SetColor("DropdownBackgroundColor", dark ? "#202832" : "#FFFFFF");
        SetColor("ShadowColor", dark ? "#091018" : "#9AA4AF");
        SetColor("SuccessColor", dark ? "#E0AC42" : "#4B5563");
        SetColor("WarningColor", dark ? "#E0AC42" : "#60758E");
        SetColor("AccentColor", dark ? "#D7A94B" : "#4E627A");
        SetColor("StatusActiveColor", dark ? "#D7A94B" : "#4E627A");
        SetColor("StatusInactiveColor", dark ? "#778390" : "#8D96A0");
        SetColor("CheckBoxCheckedBackgroundColor", dark ? "#3C2E12" : "#F8FAFC");
        SetColor("CheckBoxCheckedBorderColor", dark ? "#E0AC42" : "#4E627A");
        SetColor("CheckBoxCheckedGlyphColor", dark ? "#21170A" : "#F5F8FC");
        SetColor("StatusPillBackgroundColor", dark ? "#342814" : "#EAF0F7");
        SetColor("StatusPillBorderColor", dark ? "#6F5320" : "#BCC9D9");
        SetColor("FooterColor", dark ? "#1B232B" : "#EEF2F6");

        SetGradient(
            "BackgroundGradientBrush",
            dark
                ? ["#FF151B22", "#FF202832"]
                : ["#FFFCFDFE", "#FFF1F4F8"]);
        SetGradient(
            "TitleBarBrush",
            dark
                ? ["#F01B232B", "#F02A333E"]
                : ["#FFFDFEFF", "#FFF2F5F8"]);
        SetGradient(
            "BorderGradientBrush",
            dark
                ? ["#333B4652", "#88616E7D", "#CCD7A94B", "#88706040", "#333B4652"]
                : ["#88E8EDF4", "#BBD4DCE6", "#CC7A8899", "#AACFD7E1", "#66E8EDF4"]);
        SetGradient(
            "CardGradientBrush",
            dark
                ? ["#FF202832", "#FF26303B", "#FF2D3742", "#FF333F4B"]
                : ["#FFFFFFFF", "#FFF8FAFD", "#FFF1F4F8", "#FFEBEFF5"]);
    }

    private static void SetColor(string key, string hex)
    {
        var updated = (Color)ColorConverter.ConvertFromString(hex);
        Application.Current.Resources[key] = updated;

        var brushKey = key.Replace("Color", "Brush", StringComparison.Ordinal);
        Application.Current.Resources[brushKey] = new SolidColorBrush(updated);
    }

    private static void SetGradient(string key, string[] stops)
    {
        if (Application.Current.TryFindResource(key) is not LinearGradientBrush brush)
            return;

        var targetColors = stops.Select(static hex => (Color)ColorConverter.ConvertFromString(hex)).ToArray();
        if (brush.GradientStops.Count != targetColors.Length)
            return;

        var replacement = new LinearGradientBrush
        {
            StartPoint = brush.StartPoint,
            EndPoint = brush.EndPoint,
            MappingMode = brush.MappingMode,
            SpreadMethod = brush.SpreadMethod,
            ColorInterpolationMode = brush.ColorInterpolationMode,
            Opacity = brush.Opacity,
            Transform = brush.Transform,
            RelativeTransform = brush.RelativeTransform
        };

        for (var i = 0; i < targetColors.Length; i++)
            replacement.GradientStops.Add(new GradientStop(targetColors[i], brush.GradientStops[i].Offset));

        Application.Current.Resources[key] = replacement;
    }
}
