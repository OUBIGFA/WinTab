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
    public static string HeroDescription => Pick("一窗多标签：新窗口自动合并为当前窗口标签页", "One window, multiple tabs: auto-merge new windows into active tabs");
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
    public static string MiddleClickDescription => Pick("中键点击侧边栏、文件列表或主页中的文件夹时，新标签一律在前台打开", "Open folders middle-clicked in the navigation pane, file list or Home in the foreground instead of the background");
    public static string WheelSwitchTitle => Pick("滚轮切换标签页", "Scroll to switch tabs");
    public static string WheelSwitchDescription => Pick("鼠标在标签栏上滚动滚轮，向上切到左侧标签，向下切到右侧标签", "Scroll the mouse wheel over the tab row: up for the tab on the left, down for the tab on the right");
    public static string WheelSensitivityLabel => Pick("灵敏度", "Sensitivity");
    public static string WheelSensitivityLow => Pick("低", "Low");
    public static string WheelSensitivityMedium => Pick("中", "Medium");
    public static string WheelSensitivityHigh => Pick("高", "High");
    public static string WheelSensitivityHint(WinTab.Hooks.WheelSwitchSensitivity sensitivity) => sensitivity switch
    {
        WinTab.Hooks.WheelSwitchSensitivity.Low => Pick("滚动两格切换一次，不易误触", "Two notches per switch, hard to trigger by accident"),
        WinTab.Hooks.WheelSwitchSensitivity.High => Pick("每格立即切换，适合快速浏览多个标签", "Every notch switches at once, for skimming many tabs"),
        _ => Pick("每格切换一次，快速连滚不会连跳多个标签", "One switch per notch; a quick flick does not skip several tabs")
    };
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
    public static string TrayWheelSwitch => Pick("滚轮切换标签页", "Scroll to switch tabs");
    public static string TrayStartup => Pick("开机启动", "Start with Windows");
    public static string TrayAutoUpdate => Pick("自动检查更新", "Check for updates automatically");
    public static string TrayShowTrayIcon => Pick("显示托盘图标", "Show tray icon");
    public static string TrayCheckUpdates => Pick("检查更新", "Check for updates");
    public static string TrayExit => Pick("退出", "Exit");
}
