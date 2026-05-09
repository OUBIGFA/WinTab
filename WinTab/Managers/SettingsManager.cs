using System;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.ComponentModel;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using WinTab.Models;

namespace WinTab.Managers;

public static class SettingsManager
{
    private static readonly AppSettings Settings;
    public static event EventHandler<PropertyChangedEventArgs>? StaticPropertyChanged;

    private static readonly string SettingsFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "WinTab",
        "settings.json");

    static SettingsManager()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(SettingsFilePath)!);

        if (!File.Exists(SettingsFilePath))
        {
            Settings = new AppSettings();
            return;
        }

        try
        {
            var json = File.ReadAllText(SettingsFilePath);
            Settings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            Settings = new AppSettings();
        }
    }

    public static bool IsWindowHookActive
    {
        get => Settings.WindowHook;
        set => SetProperty(Settings.WindowHook, value, v => Settings.WindowHook = v);
    }

    public static bool ReuseTabs
    {
        get => Settings.ReuseTabs;
        set => SetProperty(Settings.ReuseTabs, value, v => Settings.ReuseTabs = v);
    }

    public static bool DoubleClickCloseTab
    {
        get => Settings.DoubleClickCloseTab;
        set => SetProperty(Settings.DoubleClickCloseTab, value, v => Settings.DoubleClickCloseTab = v);
    }

    public static bool HaveThemeIssue
    {
        get => Settings.HaveThemeIssue;
        set => SetProperty(Settings.HaveThemeIssue, value, v => Settings.HaveThemeIssue = v);
    }

    public static bool SaveClosedHistory
    {
        get => Settings.SaveClosedWindows;
        set => SetProperty(Settings.SaveClosedWindows, value, v => Settings.SaveClosedWindows = v);
    }

    public static bool RestorePreviousWindows
    {
        get => Settings.RestorePreviousWindows;
        set => SetProperty(Settings.RestorePreviousWindows, value, v => Settings.RestorePreviousWindows = v);
    }

    public static WindowRecord[]? ClosedWindows
    {
        get => Settings.ClosedWindows;
        set => SetProperty(Settings.ClosedWindows, value, v => Settings.ClosedWindows = v, notify: false);
    }

    public static bool AutoUpdate
    {
        get => Settings.AutoUpdate;
        set => SetProperty(Settings.AutoUpdate, value, v => Settings.AutoUpdate = v);
    }

    public static bool IsFirstRun
    {
        get => Settings.IsFirstRun;
        set => SetProperty(Settings.IsFirstRun, value, v => Settings.IsFirstRun = v);
    }

    public static string Language
    {
        get => string.IsNullOrWhiteSpace(Settings.Language) ? "zh-CN" : Settings.Language;
        set => SetProperty(Language, string.IsNullOrWhiteSpace(value) ? "zh-CN" : value, v => Settings.Language = v);
    }

    public static string Theme
    {
        get => string.Equals(Settings.Theme, "Dark", StringComparison.OrdinalIgnoreCase) ? "Dark" : "Light";
        set => SetProperty(Theme, string.Equals(value, "Dark", StringComparison.OrdinalIgnoreCase) ? "Dark" : "Light", v => Settings.Theme = v);
    }

    public static Size FormSize
    {
        get => Settings.FormSize;
        set => SetProperty(Settings.FormSize, value, v => Settings.FormSize = v, notify: false);
    }

    private static void SetProperty<T>(T current, T value, Action<T> assign, [CallerMemberName] string propertyName = "", bool notify = true)
    {
        if (EqualityComparer<T>.Default.Equals(current, value))
            return;

        assign(value);
        SaveSettings();

        if (notify)
            StaticPropertyChanged?.Invoke(null, new PropertyChangedEventArgs(propertyName));
    }

    public static void SaveSettings()
    {
        try
        {
            var json = JsonSerializer.Serialize(Settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SettingsFilePath, json);
        }
        catch
        {
            //
        }
    }
}

internal sealed class AppSettings
{
    public bool WindowHook { get; set; } = true;
    public bool ReuseTabs { get; set; } = true;
    public bool DoubleClickCloseTab { get; set; } = true;
    public bool HaveThemeIssue { get; set; }
    public bool SaveClosedWindows { get; set; }
    public bool RestorePreviousWindows { get; set; }
    public WindowRecord[]? ClosedWindows { get; set; }
    public bool AutoUpdate { get; set; } = true;
    public bool IsFirstRun { get; set; } = true;
    public string Language { get; set; } = "zh-CN";
    public string Theme { get; set; } = "Light";
    public Size FormSize { get; set; } = new(900, 620);
}
