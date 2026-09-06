using System.Windows;

namespace WinTab.Managers;

internal sealed record AppSettings
{
    public bool WindowHook { get; init; } = true;
    public bool ReuseTabs { get; init; } = true;
    public bool DoubleClickCloseTab { get; init; } = true;
    public bool AutoUpdate { get; init; } = true;
    public bool ShowTrayIcon { get; init; } = true;
    public string Language { get; init; } = "zh-CN";
    public string Theme { get; init; } = "Light";
    public Size FormSize { get; init; } = new(1020, 720);
}
