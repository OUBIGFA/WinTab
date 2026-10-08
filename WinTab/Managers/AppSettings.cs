using System.Globalization;
using System.Windows;

namespace WinTab.Managers;

internal sealed record AppSettings
{
    // Use Windows' display language only until the user saves an explicit choice.
    internal static string DefaultLanguage => CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "zh" ? "zh-CN" : "en-US";

    public bool WindowHook { get; init; } = true;
    public bool ReuseTabs { get; init; } = true;
    /// <summary>Only automatic reopening; capture and requested recovery do not depend on this setting.</summary>
    public bool RestoreTabs { get; init; } = false;
    public bool ReopenClosedTab { get; init; } = true;
    public bool RestoreGroupShortcutEnabled { get; init; } = true;
    public string RestoreGroupShortcut { get; init; } = "Alt+E";
    public bool ReopenTabShortcutEnabled { get; init; } = true;
    public string ReopenTabShortcut { get; init; } = "Alt+W";
    /// <summary>A shortcut needs history even when the separate recording preference is off.</summary>
    internal bool ShouldRecordClosedTabs => ReopenClosedTab || ReopenTabShortcutEnabled;
    public bool RestoreOnAnyFolder { get; init; } = false;
    /// <summary>Off: a window with a single tab is neither saved nor restored as a tab group.</summary>
    public bool RestoreSingleTab { get; init; } = false;
    public bool DoubleClickCloseTab { get; init; } = true;
    /// <summary>Whether double-click close also acts in Notepad windows, not only in Explorer.</summary>
    public bool DoubleClickCloseIncludeNotepad { get; init; } = true;
    public bool MiddleClickForegroundTab { get; init; } = true;
    public bool WheelSwitchTab { get; init; } = true;
    public string WheelSwitchSensitivity { get; init; } = "Medium";
    public bool AutoUpdate { get; init; } = true;
    public bool ShowTrayIcon { get; init; } = true;
    public string Language { get; init; } = DefaultLanguage;
    public string Theme { get; init; } = "Light";
    // Null until the user resizes the window; the initial size fits content up to the window height limit.
    public Size? FormSize { get; init; }
}
