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

    // Header / sections
    public static string HeroDescription => Pick("自动将文件夹合并为标签页", "Auto-merge folders into tabs");
    public static string ExplorerSectionTitle => Pick("资源管理器", "File Explorer");
    public static string SystemSectionTitle => Pick("系统", "System");
    public static string AboutSectionTitle => Pick("关于", "About");
    public static string BypassHint => Pick("按住 Ctrl + Shift 打开文件夹可保留独立窗口", "Hold Ctrl + Shift while opening folders to keep a separate window");

    // Settings toggles
    public static string WindowHookTitle => Pick("合并新窗口", "Merge new windows");
    public static string WindowHookDescription => Pick("新文件夹自动作为标签页打开", "Open new folders as tabs in active window");
    public static string ReuseTabsTitle => Pick("复用同一标签页", "Reuse opened tabs");
    public static string ReuseTabsDescription => Pick("重复打开同一文件夹时直接跳转", "Switch to existing tab instead of opening duplicate");
    public static string DoubleClickTitle => Pick("双击关闭标签页", "Double-click to close tabs");
    public static string DoubleClickDescription => Pick("双击标签页标题直接关闭", "Double-click any tab title to close it");
    public static string MiddleClickTitle => Pick("中键前台打开标签页", "Open middle-clicked tabs in foreground");
    public static string MiddleClickDescription => Pick("中键点击侧边栏时，新标签直接在前台打开", "Open navigation pane tabs in foreground instead of background");
    public static string StartupTitle => Pick("开机启动", "Start with Windows");
    public static string StartupDescription => Pick("开机后在后台静默运行", "Run quietly in the background on startup");
    public static string ShowTrayIconTitle => Pick("显示托盘图标", "Show tray icon");
    public static string ShowTrayIconDescription => Pick("关闭后可从右下角托盘打开", "Access from the system tray when closed");
    public static string ShowTrayIconHiddenDescription => Pick("后台继续运行，再次启动 WinTab 可打开设置", "Runs in background; launch WinTab again to open settings");
    public static string AutoUpdateTitle => Pick("自动检查更新", "Check for updates automatically");
    public static string AutoUpdateDescription => Pick("有新版本时主动提醒", "Notify when updates are available");

    // About row
    public static string CheckButton => Pick("检查更新", "Check for updates");
    public static string UpdateChecking => Pick("正在检查更新...", "Checking for updates...");
    public static string UpdateFailed => Pick("检查更新失败，请重试", "Update check failed, try again");
    public static string UpdateUpToDate => Pick("已是最新版本", "You're up to date");
    public static string UpdateNoMatchingInstaller => Pick("发现新版本，但未找到适配安装包", "New version found, but no compatible installer");
    public static string UpdateOpening => Pick("发现新版本，正在准备更新...", "Update found, opening updater...");

    // Title bar buttons
    public static string LanguageToggleTooltip => IsChinese ? "Switch to English" : "切换到中文";
    public static string ThemeToggleTooltip(bool isDarkTheme) =>
        isDarkTheme ? Pick("切换为浅色", "Switch to light mode") : Pick("切换为深色", "Switch to dark mode");

    // Tray icon
    public static string TrayTooltip => Pick("WinTab - 资源管理器标签页工具", "WinTab - File Explorer Tab Utility");
    public static string TrayOpen => Pick("打开 WinTab", "Open WinTab");
    public static string TrayWindowHook => Pick("合并新窗口", "Merge new windows");
    public static string TrayReuseTabs => Pick("复用同一标签页", "Reuse opened tabs");
    public static string TrayDoubleClickClose => Pick("双击关闭标签页", "Double-click to close tabs");
    public static string TrayMiddleClickForeground => Pick("中键前台打开标签页", "Open middle-clicked tabs in foreground");
    public static string TrayStartup => Pick("开机启动", "Start with Windows");
    public static string TrayAutoUpdate => Pick("自动检查更新", "Check for updates automatically");
    public static string TrayShowTrayIcon => Pick("显示托盘图标", "Show tray icon");
    public static string TrayCheckUpdates => Pick("检查更新", "Check for updates");
    public static string TrayExit => Pick("退出", "Exit");
}
