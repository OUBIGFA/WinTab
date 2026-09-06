using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using WinTab.Helpers;

namespace WinTab.Managers;

public static class SettingsManager
{
    private static readonly SettingsStore Store = new(Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "WinTab", Constants.SettingsFileName));

    public static event EventHandler<PropertyChangedEventArgs>? StaticPropertyChanged;
    public static event Action? StorageErrorChanged;
    public static Exception? StorageError => Store.LastError;

    static SettingsManager()
    {
        Store.ErrorChanged += () => StorageErrorChanged?.Invoke();
    }

    public static bool IsWindowHookActive
    {
        get => Store.Snapshot.WindowHook;
        set => SetProperty(settings => settings with { WindowHook = value });
    }

    public static bool ReuseTabs
    {
        get => Store.Snapshot.ReuseTabs;
        set => SetProperty(settings => settings with { ReuseTabs = value });
    }

    public static bool DoubleClickCloseTab
    {
        get => Store.Snapshot.DoubleClickCloseTab;
        set => SetProperty(settings => settings with { DoubleClickCloseTab = value });
    }

    public static bool AutoUpdate
    {
        get => Store.Snapshot.AutoUpdate;
        set => SetProperty(settings => settings with { AutoUpdate = value });
    }

    public static bool ShowTrayIcon
    {
        get => Store.Snapshot.ShowTrayIcon;
        set => SetProperty(settings => settings with { ShowTrayIcon = value });
    }

    public static string Language
    {
        get => NormalizeLanguage(Store.Snapshot.Language);
        set => SetProperty(settings => settings with { Language = NormalizeLanguage(value) });
    }

    public static string Theme
    {
        get => NormalizeTheme(Store.Snapshot.Theme);
        set => SetProperty(settings => settings with { Theme = NormalizeTheme(value) });
    }

    public static Size FormSize
    {
        get => Store.Snapshot.FormSize;
        set => SetProperty(settings => settings with { FormSize = value }, notify: false, deferred: true);
    }

    private static void SetProperty(Func<AppSettings, AppSettings> update,
        [CallerMemberName] string propertyName = "", bool notify = true, bool deferred = false)
    {
        if (Store.Update(update, deferred) && notify)
            StaticPropertyChanged?.Invoke(null, new PropertyChangedEventArgs(propertyName));
    }

    public static void SaveSettings() => _ = Store.FlushAsync();

    public static async Task<bool> FlushSettingsAsync(TimeSpan timeout)
    {
        try
        {
            return await Store.FlushAsync().WaitAsync(timeout).ConfigureAwait(false);
        }
        catch (TimeoutException)
        {
            System.Diagnostics.Trace.TraceError("Settings save timed out; the last valid settings file is retained.");
            return false;
        }
    }

    private static string NormalizeLanguage(string? value) => string.IsNullOrWhiteSpace(value) ? "zh-CN" : value;

    private static string NormalizeTheme(string? value) =>
        string.Equals(value, "Dark", StringComparison.OrdinalIgnoreCase) ? "Dark" : "Light";
}
