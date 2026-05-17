using System;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.ComponentModel;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Threading;
using WinTab.Models;

namespace WinTab.Managers;

public static class SettingsManager
{
    private static readonly AppSettings Settings;
    private static readonly object SettingsLock = new();
    private static readonly JsonSerializerOptions SerializerOptions = new() { WriteIndented = true };
    private static readonly TimeSpan DeferredSaveDelay = TimeSpan.FromMilliseconds(500);
    private static Timer? _deferredSaveTimer;
    private static bool _hasPendingDeferredSave;
    public static event EventHandler<PropertyChangedEventArgs>? StaticPropertyChanged;

    private static readonly string SettingsDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "WinTab");
    private static readonly string SettingsFilePath = Path.Combine(SettingsDirectory, "settings.json");

    static SettingsManager()
    {
        Directory.CreateDirectory(SettingsDirectory);

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
        get => ReadProperty(() => Settings.WindowHook);
        set => SetProperty(() => Settings.WindowHook, value, v => Settings.WindowHook = v);
    }

    public static bool ReuseTabs
    {
        get => ReadProperty(() => Settings.ReuseTabs);
        set => SetProperty(() => Settings.ReuseTabs, value, v => Settings.ReuseTabs = v);
    }

    public static bool DoubleClickCloseTab
    {
        get => ReadProperty(() => Settings.DoubleClickCloseTab);
        set => SetProperty(() => Settings.DoubleClickCloseTab, value, v => Settings.DoubleClickCloseTab = v);
    }

    public static bool HaveThemeIssue
    {
        get => ReadProperty(() => Settings.HaveThemeIssue);
        set => SetProperty(() => Settings.HaveThemeIssue, value, v => Settings.HaveThemeIssue = v);
    }

    public static bool SaveClosedHistory
    {
        get => ReadProperty(() => Settings.SaveClosedWindows);
        set => SetProperty(() => Settings.SaveClosedWindows, value, v => Settings.SaveClosedWindows = v);
    }

    public static bool RestorePreviousWindows
    {
        get => ReadProperty(() => Settings.RestorePreviousWindows);
        set => SetProperty(() => Settings.RestorePreviousWindows, value, v => Settings.RestorePreviousWindows = v);
    }

    public static WindowRecord[]? ClosedWindows
    {
        get => ReadProperty(() => Settings.ClosedWindows);
        set => SetProperty(() => Settings.ClosedWindows, value, v => Settings.ClosedWindows = v, notify: false);
    }

    public static bool AutoUpdate
    {
        get => ReadProperty(() => Settings.AutoUpdate);
        set => SetProperty(() => Settings.AutoUpdate, value, v => Settings.AutoUpdate = v);
    }

    public static bool ShowTrayIcon
    {
        get => ReadProperty(() => Settings.ShowTrayIcon);
        set => SetProperty(() => Settings.ShowTrayIcon, value, v => Settings.ShowTrayIcon = v);
    }

    public static bool IsFirstRun
    {
        get => ReadProperty(() => Settings.IsFirstRun);
        set => SetProperty(() => Settings.IsFirstRun, value, v => Settings.IsFirstRun = v);
    }

    public static string Language
    {
        get => ReadProperty(() => NormalizeLanguage(Settings.Language));
        set => SetProperty(() => NormalizeLanguage(Settings.Language), NormalizeLanguage(value), v => Settings.Language = v);
    }

    public static string Theme
    {
        get => ReadProperty(() => NormalizeTheme(Settings.Theme));
        set => SetProperty(() => NormalizeTheme(Settings.Theme), NormalizeTheme(value), v => Settings.Theme = v);
    }

    public static Size FormSize
    {
        get => ReadProperty(() => Settings.FormSize);
        set => SetProperty(() => Settings.FormSize, value, v => Settings.FormSize = v, notify: false, saveMode: SaveMode.Deferred);
    }

    private static T ReadProperty<T>(Func<T> read)
    {
        lock (SettingsLock)
            return read();
    }

    private static void SetProperty<T>(
        Func<T> readCurrent,
        T value,
        Action<T> assign,
        [CallerMemberName] string propertyName = "",
        bool notify = true,
        SaveMode saveMode = SaveMode.Immediate)
    {
        var changed = false;
        lock (SettingsLock)
        {
            if (EqualityComparer<T>.Default.Equals(readCurrent(), value))
                return;

            assign(value);
            if (saveMode == SaveMode.Immediate)
            {
                SaveSettingsCore();
                CancelDeferredSaveCore();
            }
            else
            {
                ScheduleDeferredSaveCore();
            }
            changed = true;
        }

        if (changed && notify)
            StaticPropertyChanged?.Invoke(null, new PropertyChangedEventArgs(propertyName));
    }

    public static void SaveSettings()
    {
        lock (SettingsLock)
        {
            CancelDeferredSaveCore();
            SaveSettingsCore();
        }
    }

    private static void SaveSettingsCore()
    {
        try
        {
            Directory.CreateDirectory(SettingsDirectory);
            var json = JsonSerializer.Serialize(Settings, SerializerOptions);
            File.WriteAllText(SettingsFilePath, json);
        }
        catch
        {
            //
        }
    }

    private static void ScheduleDeferredSaveCore()
    {
        _hasPendingDeferredSave = true;
        _deferredSaveTimer ??= new Timer(_ => FlushDeferredSave(), null, Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
        _deferredSaveTimer.Change(DeferredSaveDelay, Timeout.InfiniteTimeSpan);
    }

    private static void FlushDeferredSave()
    {
        lock (SettingsLock)
        {
            if (!_hasPendingDeferredSave)
                return;

            _hasPendingDeferredSave = false;
            SaveSettingsCore();
        }
    }

    private static void CancelDeferredSaveCore()
    {
        _hasPendingDeferredSave = false;
        _deferredSaveTimer?.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
    }

    private static string NormalizeLanguage(string? value) => string.IsNullOrWhiteSpace(value) ? "zh-CN" : value;

    private static string NormalizeTheme(string? value) =>
        string.Equals(value, "Dark", StringComparison.OrdinalIgnoreCase) ? "Dark" : "Light";

    private enum SaveMode
    {
        Immediate,
        Deferred
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
    public bool ShowTrayIcon { get; set; } = true;
    public bool IsFirstRun { get; set; } = true;
    public string Language { get; set; } = "zh-CN";
    public string Theme { get; set; } = "Light";
    public Size FormSize { get; set; } = new(1020, 720);
}
