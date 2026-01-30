using System.Globalization;
using System.Linq;
using System.Windows;
using WinTab.Core;

namespace WinTab.App.Localization;

public static class LocalizationManager
{
    private static readonly ResourceDictionary EnglishResources = new();
    private static readonly ResourceDictionary ChineseResources = new();
    private static bool _isInitialized = false;

    public static void Initialize()
    {
        if (_isInitialized)
            return;

        // Load resource dictionaries
        EnglishResources.Source = new Uri("pack://application:,,,/WinTab.App;component/Localization/Strings.en-US.xaml");
        ChineseResources.Source = new Uri("pack://application:,,,/WinTab.App;component/Localization/Strings.zh-CN.xaml");
        
        _isInitialized = true;
    }

    public static void ApplyLanguage(Language language)
    {
        if (!_isInitialized)
            Initialize();

        var currentDict = System.Windows.Application.Current.Resources.MergedDictionaries.FirstOrDefault(d => 
            d.Source?.ToString().Contains("Strings.") == true);

        if (currentDict != null)
        {
            System.Windows.Application.Current.Resources.MergedDictionaries.Remove(currentDict);
        }

        switch (language)
        {
            case Language.Chinese:
                System.Windows.Application.Current.Resources.MergedDictionaries.Add(ChineseResources);
                CultureInfo.CurrentCulture = new CultureInfo("zh-CN");
                CultureInfo.CurrentUICulture = new CultureInfo("zh-CN");
                break;
            case Language.English:
            default:
                System.Windows.Application.Current.Resources.MergedDictionaries.Add(EnglishResources);
                CultureInfo.CurrentCulture = new CultureInfo("en-US");
                CultureInfo.CurrentUICulture = new CultureInfo("en-US");
                break;
        }
    }

    public static string GetString(string key)
    {
        if (!_isInitialized)
            Initialize();

        // Try to get from current resources
        var currentDict = System.Windows.Application.Current.Resources.MergedDictionaries.FirstOrDefault(d => 
            d.Source?.ToString().Contains("Strings.") == true);

        if (currentDict != null && currentDict.Contains(key))
        {
            return currentDict[key]?.ToString() ?? key;
        }

        // Fallback to English
        if (EnglishResources.Contains(key))
        {
            return EnglishResources[key]?.ToString() ?? key;
        }

        return key;
    }
}
