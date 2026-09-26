using System.Windows;

namespace WinTab.Managers;

internal sealed record AppSettings
{
    public bool WindowHook { get; init; } = true;
    public bool ReuseTabs { get; init; } = true;
    public bool RestoreTabs { get; init; } = false;
    public bool RestoreOnAnyFolder { get; init; } = false;
    /// <summary>Off: a window with a single tab is neither saved nor restored as a tab group.</summary>
    public bool RestoreSingleTab { get; init; } = false;
    public bool DoubleClickCloseTab { get; init; } = true;
    public bool MiddleClickForegroundTab { get; init; } = true;
    public bool WheelSwitchTab { get; init; } = true;
    public string WheelSwitchSensitivity { get; init; } = "Medium";
    public bool AutoUpdate { get; init; } = true;
    public bool ShowTrayIcon { get; init; } = true;
    public string Language { get; init; } = "zh-CN";
    public string Theme { get; init; } = "Light";
    // Null until the user resizes the window; the window then opens tall enough to fit its content.
    public Size? FormSize { get; init; }
}
