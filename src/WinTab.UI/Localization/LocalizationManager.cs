using System.Globalization;
using System.Windows;
using WinTab.Core.Enums;

namespace WinTab.UI.Localization;

public static class LocalizationManager
{
    private static readonly object Lock = new();
    private static ResourceDictionary? _chineseResources;
    private static ResourceDictionary? _englishResources;
    private static Language _currentLanguage = Language.Chinese;

    public static Language CurrentLanguage => _currentLanguage;

    public static void ApplyLanguage(Language language)
    {
        lock (Lock)
        {
            EnsureResourcesLoaded();

            var app = Application.Current;
            if (app is null) return;

            // Remove existing localization dictionary
            var toRemove = app.Resources.MergedDictionaries
                .FirstOrDefault(d => d.Source?.ToString().Contains("Strings.") == true);
            if (toRemove is not null)
                app.Resources.MergedDictionaries.Remove(toRemove);

            // Add the selected language
            var dict = language switch
            {
                Language.English => _englishResources!,
                _ => _chineseResources!
            };

            app.Resources.MergedDictionaries.Add(dict);
            _currentLanguage = language;

            // Update culture
            var culture = language switch
            {
                Language.English => new CultureInfo("en-US"),
                _ => new CultureInfo("zh-CN")
            };

            CultureInfo.CurrentCulture = culture;
            CultureInfo.CurrentUICulture = culture;
        }
    }

    public static string GetString(string key)
    {
        lock (Lock)
        {
            EnsureResourcesLoaded();

            var app = Application.Current;
            if (app is null) return key;

            // Try current merged dictionaries first
            if (app.Resources[key] is string value)
                return value;

            // Fallback to English
            if (_englishResources?.Contains(key) == true)
                return _englishResources[key]?.ToString() ?? key;

            return key;
        }
    }

    private static void EnsureResourcesLoaded()
    {
        if (_chineseResources is not null) return;

        _chineseResources = new ResourceDictionary
        {
            Source = new Uri("pack://application:,,,/WinTab.UI;component/Localization/Strings.zh-CN.xaml")
        };
        _englishResources = new ResourceDictionary
        {
            Source = new Uri("pack://application:,,,/WinTab.UI;component/Localization/Strings.en-US.xaml")
        };
    }
}
