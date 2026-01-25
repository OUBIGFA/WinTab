using System;
using System.Globalization;
using System.Threading;
using System.Windows;
using WinTab.Managers;

namespace WinTab.UI.Localization;

public static class LocalizationManager
{
    private const string DefaultLanguage = "en-US";

    public static string CurrentLanguage => SettingsManager.Language;

    public static void Initialize()
    {
        var language = SettingsManager.Language;
        if (string.IsNullOrWhiteSpace(language))
        {
            language = GetDefaultLanguageFromSystem();
            SettingsManager.Language = language;
        }

        ApplyLanguage(language);
    }

    public static void ChangeLanguage(string language)
    {
        if (string.IsNullOrWhiteSpace(language))
            language = DefaultLanguage;

        SettingsManager.Language = language;
        ApplyLanguage(language);
    }

    private static string GetDefaultLanguageFromSystem()
    {
        var culture = CultureInfo.CurrentUICulture;
        if (culture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase))
            return "zh-CN";

        return DefaultLanguage;
    }

    private static void ApplyLanguage(string language)
    {
        var uri = new Uri($"pack://application:,,,/WinTab;component/UI/Localization/Strings.{language}.xaml", UriKind.Absolute);

        var dictionary = new ResourceDictionary { Source = uri };

        var appResources = Application.Current.Resources;
        RemoveExistingLocalizationDictionaries(appResources);
        appResources.MergedDictionaries.Add(dictionary);

        SetThreadCulture(language);
    }

    private static void RemoveExistingLocalizationDictionaries(ResourceDictionary appResources)
    {
        for (var i = appResources.MergedDictionaries.Count - 1; i >= 0; i--)
        {
            var source = appResources.MergedDictionaries[i].Source;
            if (source == null) continue;

            var path = source.ToString();
            if (path.Contains("/UI/Localization/Strings.", StringComparison.OrdinalIgnoreCase))
                appResources.MergedDictionaries.RemoveAt(i);
        }
    }

    private static void SetThreadCulture(string language)
    {
        try
        {
            var culture = new CultureInfo(language);
            Thread.CurrentThread.CurrentCulture = culture;
            Thread.CurrentThread.CurrentUICulture = culture;
        }
        catch
        {
            var fallback = CultureInfo.InvariantCulture;
            Thread.CurrentThread.CurrentCulture = fallback;
            Thread.CurrentThread.CurrentUICulture = fallback;
        }
    }
}
