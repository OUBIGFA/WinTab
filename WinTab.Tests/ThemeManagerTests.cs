using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using WinTab.Helpers;
using WinTab.Managers;

internal static class ThemeManagerTests
{
    private static readonly string[] PaletteKeys =
    {
        "PrimaryBrush",
        "PrimaryLightBrush",
        "TextPrimaryBrush",
        "TextSecondaryBrush",
        "TextTertiaryBrush",
        "TextAccentBrush",
        "BorderBrush",
        "DividerBrush",
        "ControlBackgroundBrush",
        "ControlHoverBrush",
        "DropdownBackgroundBrush",
        "ShadowColor",
        "CheckBoxCheckedBackgroundBrush",
        "CheckBoxCheckedBorderBrush",
        "CheckBoxCheckedGlyphBrush",
        "StatusPillBackgroundBrush",
        "StatusPillBorderBrush",
        "FocusRingBrush",
        "WindowBackgroundBrush",
        "WindowTitleBarBrush",
        "WindowBorderBrush",
        "SurfaceBrush",
        "SurfaceMutedBrush",
        "SurfaceRaisedBrush"
    };

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("Theme palettes keep neutral gray hierarchy and readable contrast", ThemePalettesAreBalanced);
    }

    private static async Task ThemePalettesAreBalanced()
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(() =>
        {
            _ = Application.Current ?? new Application();

            ThemeManager.ApplyTheme(dark: false);
            AssertPaletteIsNeutral();
            AssertNeutralGray(GetColor("WindowBackgroundBrush"), "Light window background");
            AssertNeutralGray(GetColor("SurfaceMutedBrush"), "Light muted surface");
            Check.That(GetGrayLevel("WindowBackgroundBrush") - GetGrayLevel("SurfaceMutedBrush") >= 8,
                "Light surfaces need a visible tonal separation.");
            AssertContrast(GetColor("TextPrimaryBrush"), GetColor("WindowBackgroundBrush"), 7.0,
                "Light primary text");
            AssertContrast(GetColor("TextSecondaryBrush"), GetColor("SurfaceRaisedBrush"), 4.5,
                "Light secondary text");

            ThemeManager.ApplyTheme(dark: true);
            AssertPaletteIsNeutral();
            AssertContrast(GetColor("TextPrimaryBrush"), GetColor("WindowBackgroundBrush"), 7.0,
                "Dark primary text");
            AssertContrast(GetColor("TextSecondaryBrush"), GetColor("SurfaceRaisedBrush"), 4.5,
                "Dark secondary text");
        }, CancellationToken.None, TaskCreationOptions.None, scheduler);
    }

    private static Color GetColor(string key)
    {
        var resource = Application.Current.Resources[key];
        return resource switch
        {
            SolidColorBrush brush => brush.Color,
            Color color => color,
            _ => throw new InvalidOperationException($"Theme resource '{key}' is not a color or brush.")
        };
    }

    private static byte GetGrayLevel(string key) => GetColor(key).R;

    private static void AssertPaletteIsNeutral()
    {
        foreach (var key in PaletteKeys)
        {
            var color = GetColor(key);
            AssertNeutralGray(color, key);
        }
    }

    private static void AssertNeutralGray(Color color, string name)
    {
        Check.Equal(color.R, color.G, $"{name} should remain neutral gray.");
        Check.Equal(color.G, color.B, $"{name} should remain neutral gray.");
        Check.That(color.R is > 0 and < 255, $"{name} should avoid dead black or dead white.");
    }

    private static void AssertContrast(Color foreground, Color background, double minimum, string name)
    {
        var lighter = Math.Max(RelativeLuminance(foreground), RelativeLuminance(background));
        var darker = Math.Min(RelativeLuminance(foreground), RelativeLuminance(background));
        var ratio = (lighter + 0.05) / (darker + 0.05);
        Check.That(ratio >= minimum, $"{name} contrast {ratio:F2}:1 must be at least {minimum:F1}:1.");
    }

    private static double RelativeLuminance(Color color)
    {
        static double Linearize(byte channel)
        {
            var value = channel / 255.0;
            return value <= 0.04045 ? value / 12.92 : Math.Pow((value + 0.055) / 1.055, 2.4);
        }

        return (0.2126 * Linearize(color.R)) +
               (0.7152 * Linearize(color.G)) +
               (0.0722 * Linearize(color.B));
    }
}
