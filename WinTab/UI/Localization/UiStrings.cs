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
    public static string RestoreTabsTitle => Pick("自动恢复上次标签组", "Auto-restore last tab group");
    public static string RestoreTabsDescription => Pick("无其他资源管理器窗口时，下次打开自动恢复已保存的标签组", "Automatically restore the saved group on launch when no other Explorer windows are open");
    public static string RestoreSingleTabTitle => Pick("恢复单个标签", "Restore single-tab windows");
    public static string RestoreSingleTabDescription => Pick("默认关闭；关闭时仅保存和恢复含两个及以上标签的窗口", "Off by default; only save and restore windows with two or more tabs when off");
    public static string RestoreNormalLaunchOnly => Pick("仅普通启动", "Normal launch only");
    public static string RestoreAnyFolder => Pick("打开任意文件夹", "Any folder");
    public static string RestoreModeHint(bool anyFolder) => anyFolder
        ? Pick("保留最初打开的标签并停留在它上面，包括主页／此电脑等启动页", "Keep the initial tab active, including start pages such as Home or This PC")
        : Pick("仅从主页／此电脑等默认位置普通启动时恢复；打开特定文件夹不恢复", "Restore only on a normal launch to a default location such as Home or This PC, not when opening a specific folder");
    public static string RestoreExclusions => Pick("拖出的窗口和 Ctrl + Shift 独立窗口不触发恢复", "Torn-off windows and Ctrl + Shift separate windows do not trigger restoration");
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

    // Manual recovery, independent of automatic restoration.
    public static string RecoveryTitle => Pick("标签恢复", "Tab recovery");
    public static string RecoveryDescription => Pick("后台持续保存标签组；关机或资源管理器崩溃后保留最后使用的窗口", "Groups are saved continuously; recover the last-used window after shutdown or an Explorer crash");
    public static string RestoreGroupCommand => Pick("恢复上次标签组", "Restore last tab group");
    public static string ReopenTabCommand => Pick("恢复上一个标签页", "Reopen last closed tab");
    public static string RecordClosedTabs => Pick("记住最近关闭的标签页", "Remember recently closed tabs");
    public static string RestoreNow => Pick("立即恢复", "Restore now");
    public static string ShortcutEnabled => Pick("启用快捷键", "Enable shortcut");
    public static string ShortcutSave => Pick("应用快捷键", "Apply shortcuts");
    public static string ShortcutHint => Pick("快捷键仅在资源管理器中生效，不影响浏览器；支持 Ctrl / Alt 与字母、数字、F1–F24，可加 Shift", "Shortcuts work only in Explorer, not browsers; use Ctrl / Alt with a letter, digit or F1–F24; Shift is optional");
    public static string ShortcutInvalid => Pick("快捷键无效，请使用 Ctrl / Alt + 字母、数字或 F1–F24", "Invalid shortcut; use Ctrl / Alt + a letter, digit or F1–F24");
    public static string ShortcutDuplicate => Pick("两个恢复功能不能使用相同快捷键", "The two recovery commands must use different shortcuts");
    public static string ShortcutUnavailable => Pick("快捷键未能启用；仍可使用恢复按钮", "Shortcuts could not be enabled; recovery buttons remain available");
    public static string ShortcutSaved => Pick("快捷键已应用", "Shortcuts applied");
    public static string SessionResult(WinTab.Hooks.SessionCommandResult result) => result switch
    {
        WinTab.Hooks.SessionCommandResult.Completed => Pick("恢复完成", "Restored"),
        WinTab.Hooks.SessionCommandResult.CompletedWithSkips => Pick("恢复完成；已跳过不存在或不支持的位置", "Restored; missing or unsupported locations were skipped"),
        WinTab.Hooks.SessionCommandResult.NothingSaved => Pick("尚无已保存的标签组；默认仅记录含两个及以上标签的窗口", "No saved group yet; by default, groups need at least two tabs"),
        WinTab.Hooks.SessionCommandResult.NothingAvailable => Pick("记录中的位置已不存在或暂不支持恢复（例如网络位置）", "Recorded locations are missing or unsupported (such as network locations)"),
        WinTab.Hooks.SessionCommandResult.NoClosedTab => Pick("没有可恢复的最近关闭标签页", "No recently closed tab to reopen"),
        WinTab.Hooks.SessionCommandResult.Busy => Pick("正在恢复，请稍后再试", "A restore is in progress; try again shortly"),
        WinTab.Hooks.SessionCommandResult.NotReady => Pick("正在连接资源管理器，请稍后重试", "Connecting to Explorer; try again shortly"),
        _ => Pick("恢复未完成；请检查现有标签页后重试", "Restore did not finish; check the existing tabs before retrying")
    };

    // About row
    public static string CheckButton => Pick("检查更新", "Check for updates");
    public static string OpenLogsButton => Pick("打开日志", "Open logs");
    public static string OpenLogsFailed => Pick("无法打开日志文件夹", "Could not open the log folder");
    public static string UpdateChecking => Pick("正在检查更新", "Checking for updates");
    public static string UpdateFailed => Pick("检查更新失败，请重试", "Update check failed, try again");
    public static string UpdateUpToDate => Pick("已是最新版本", "You're up to date");
    public static string UpdateNoMatchingInstaller => Pick("发现新版本，但未找到适配安装包", "New version found, but no compatible installer");
    public static string UpdateOpening => Pick("发现新版本，正在准备更新", "Update found, opening updater");

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
