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

        SetColor("PrimaryColor", dark ? "#7CA7D8" : "#3F4A56");
        SetColor("PrimaryLightColor", dark ? "#A8C4E7" : "#5C6875");
        SetColor("TextPrimaryColor", dark ? "#EEF2F6" : "#20262D");
        SetColor("TextSecondaryColor", dark ? "#AEB8C4" : "#5B6672");
        SetColor("TextTertiaryColor", dark ? "#84909C" : "#78828D");
        SetColor("TextAccentColor", dark ? "#B7D3F4" : "#334155");
        SetColor("BorderColor", dark ? "#3C4652" : "#D2D8DF");
        SetColor("ControlBackgroundColor", dark ? "#202832" : "#F6F7F8");
        SetColor("ControlHoverColor", dark ? "#2B3541" : "#ECEFF2");
        SetColor("DangerColor", dark ? "#F06A6A" : "#C93F3F");
        SetColor("DropdownBackgroundColor", dark ? "#202832" : "#FFFFFF");
        SetColor("ShadowColor", dark ? "#091018" : "#9AA4AF");
        SetColor("SuccessColor", dark ? "#87B7F0" : "#4B5563");
        SetColor("WarningColor", dark ? "#E0AC42" : "#B57B16");
        SetColor("AccentColor", dark ? "#D7A94B" : "#4B5563");
        SetColor("StatusActiveColor", dark ? "#87B7F0" : "#4B5563");
        SetColor("StatusInactiveColor", dark ? "#778390" : "#8D96A0");
        SetColor("CheckBoxCheckedBackgroundColor", dark ? "#334155" : "#F8FAFC");
        SetColor("CheckBoxCheckedBorderColor", dark ? "#7CA7D8" : "#64748B");
        SetColor("CheckBoxCheckedGlyphColor", dark ? "#EEF2F6" : "#20262D");

        SetGradient(
            "BackgroundGradientBrush",
            dark
                ? ["#FF151B22", "#FF202832"]
                : ["#FFFBFBFA", "#FFF0F2F4"]);
        SetGradient(
            "TitleBarBrush",
            dark
                ? ["#F01B232B", "#F02A333E"]
                : ["#FFF8F8F7", "#FFECEFF1"]);
        SetGradient(
            "BorderGradientBrush",
            dark
                ? ["#333B4652", "#88616E7D", "#CC7CA7D8", "#88706040", "#333B4652"]
                : ["#66D8DDE3", "#AAB9C0C9", "#CC6F7A86", "#AAAEB6BF", "#66D8DDE3"]);
        SetGradient(
            "CardGradientBrush",
            dark
                ? ["#FF202832", "#FF26303B", "#FF2D3742", "#FF333F4B"]
                : ["#FFFFFFFF", "#FFF9FAFA", "#FFF3F4F5", "#FFEFF1F3"]);
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
