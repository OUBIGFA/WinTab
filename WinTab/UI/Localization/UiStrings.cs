using System;
using WinTab.Managers;

namespace WinTab.UI.Localization;

/// <summary>
/// The single bilingual string table for the settings window and the tray menu.
/// Every entry is (Simplified Chinese, English); the active language comes from settings.
/// </summary>
internal static class UiStrings
{
    public static bool IsChinese => string.Equals(SettingsManager.Language, "zh-CN", StringComparison.OrdinalIgnoreCase);

    public static string Pick(string zh, string en) => IsChinese ? zh : en;

    // Hero / status
    public static string HeroDescription => Pick("让资源管理器窗口回到同一组标签", "Keep File Explorer windows in one tab set");
    public static string StatusRunning => Pick("运行中", "Running");
    public static string StatusTrayAvailable => Pick("可从托盘打开", "Available from the tray");
    public static string StatusTrayHidden => Pick("托盘图标已隐藏", "Tray icon is hidden");
    public static string StatusBypassHint => Pick("按住 Ctrl + Shift 可打开独立窗口", "Hold Ctrl + Shift to open a separate window");
    public static string Version(string version) => Pick($"版本：v{version}", $"Version: v{version}");

    // Settings toggles
    public static string WindowHookTitle => Pick("合并新窗口", "Merge new windows");
    public static string WindowHookDescription => Pick("将新打开的文件夹收回当前标签组", "Send new folders back to the active Explorer tab group");
    public static string ReuseTabsTitle => Pick("复用已有标签", "Reuse existing tabs");
    public static string ReuseTabsDescription => Pick("路径已打开时，直接聚焦对应标签", "Open a matching path by focusing its current tab");
    public static string DoubleClickTitle => Pick("双击关闭标签", "Double-click closes tab");
    public static string DoubleClickDescription => Pick("在标题区双击关闭当前标签", "Close the current Explorer tab from its title area");
    public static string StartupTitle => Pick("开机启动", "Start with Windows");
    public static string StartupDescription => Pick("登录后静默运行", "Run quietly after sign-in");
    public static string ShowTrayIconTitle => Pick("显示托盘图标", "Show tray icon");
    public static string ShowTrayIconDescription => Pick("关闭窗口后可从通知区打开", "Keep WinTab available from the notification area");
    public static string AutoUpdateTitle => Pick("检查更新", "Check for updates");
    public static string AutoUpdateDescription => Pick("有 GitHub Release 时提示", "Notify when a GitHub release is available");

    // Maintenance area
    public static string MaintenanceTitle => Pick("维护", "Maintenance");
    public static string MaintenanceTrayVisible => Pick("关闭此窗口后，WinTab 继续留在托盘", "The app stays in the tray when this window closes");
    public static string MaintenanceTrayHidden => Pick("托盘图标隐藏时，WinTab 会直接在后台运行", "When the tray icon is hidden, WinTab keeps running in the background");
    public static string CheckButton => Pick("检查", "Check");
    public static string CheckingButton => Pick("检查中", "Checking");
    public static string HideButton => Pick("隐藏", "Hide");
    public static string UpdateChecking => Pick("正在联网检查最新版本", "Checking the latest release online");
    public static string UpdateFailed => Pick("检查失败，稍后再试", "Update check failed, try again later");
    public static string UpdateUpToDate => Pick("当前已是最新版本", "You're on the latest version");
    public static string UpdateNoMatchingInstaller => Pick("发现新版本，但未找到匹配当前架构的安装器", "Update found, but no installer matches this device");
    public static string UpdateOpening => Pick("发现新版本，正在打开更新窗口", "Update found, opening the update window");

    // Title bar buttons
    public static string LanguageToggleTooltip => IsChinese ? "Switch to English" : "切换到中文";
    public static string ThemeToggleTooltip(bool isDarkTheme) =>
        isDarkTheme ? Pick("切换为浅色", "Switch to light mode") : Pick("切换为深色", "Switch to dark mode");

    // Tray icon
    public static string TrayTooltip => Pick("WinTab 会将文件夹尽量保持在资源管理器标签页中", "WinTab keeps File Explorer folders in tabs");
    public static string TrayOpen => Pick("打开 WinTab", "Open WinTab");
    public static string TrayWindowHook => Pick("自动合并资源管理器窗口", "Auto-merge Explorer windows");
    public static string TrayReuseTabs => Pick("复用标签", "Reuse tabs");
    public static string TrayDoubleClickClose => Pick("双击关闭标签页", "Double-click to close tab");
    public static string TrayStartup => Pick("开机启动", "Start with Windows");
    public static string TrayAutoUpdate => Pick("自动更新", "Automatic updates");
    public static string TrayShowTrayIcon => Pick("显示托盘图标", "Show tray icon");
    public static string TrayCheckUpdates => Pick("检查更新", "Check for updates");
    public static string TrayExit => Pick("退出", "Exit");
}
