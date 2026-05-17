using System;
using System.Windows;
using System.Windows.Controls;
using WinTab.Helpers;
using WinTab.Managers;

namespace WinTab.UI.Views.Controls;

public partial class SystemTrayIcon : UserControl, IDisposable
{
    private readonly HookManager _hookManager;
    private readonly Action _showWindowAction;
    private readonly Action _exitAction;
    private bool _disposed;

    public event EventHandler? SettingsChanged;

    public SystemTrayIcon(HookManager hookManager, Action showWindowAction, Action exitAction)
    {
        InitializeComponent();

        _hookManager = hookManager;
        _showWindowAction = showWindowAction;
        _exitAction = exitAction;

        TrayIcon.Icon = Helper.GetIcon();
        ApplyLanguage(string.Equals(SettingsManager.Language, "zh-CN", StringComparison.OrdinalIgnoreCase));
        ApplyThemeText(ThemeManager.IsDarkTheme, string.Equals(SettingsManager.Language, "zh-CN", StringComparison.OrdinalIgnoreCase));

        OpenSettingsMenu.Click += (_, _) => _showWindowAction();
        WindowHookMenu.Click += WindowHookMenu_Click;
        ReuseTabsMenu.Click += ReuseTabsMenu_Click;
        DoubleClickCloseMenu.Click += DoubleClickCloseMenu_Click;
        StartupMenu.Click += StartupMenu_Click;
        AutoUpdateMenu.Click += AutoUpdateMenu_Click;
        ShowTrayIconMenu.Click += ShowTrayIconMenu_Click;
        CheckUpdatesMenu.Click += (_, _) => UpdateManager.CheckForUpdates();
        ExitMenu.Click += (_, _) => _exitAction();

        _hookManager.StateChanged += RefreshState;
        _hookManager.ShellInitialized += RefreshState;
        RefreshState();
    }

    public void ApplyLanguage(bool zh)
    {
        TrayIcon.ToolTipText = zh
            ? "WinTab \u4f1a\u5c06\u6587\u4ef6\u5939\u5c3d\u91cf\u4fdd\u6301\u5728\u8d44\u6e90\u7ba1\u7406\u5668\u6807\u7b7e\u9875\u4e2d\u3002"
            : Constants.NotifyIconText;
        OpenSettingsMenu.Header = zh ? "\u6253\u5f00 WinTab" : "Open WinTab";
        WindowHookMenu.Header = zh ? "\u81ea\u52a8\u5408\u5e76\u8d44\u6e90\u7ba1\u7406\u5668\u7a97\u53e3" : "Auto-merge Explorer windows";
        ReuseTabsMenu.Header = zh ? "\u590d\u7528\u6807\u7b7e" : "Reuse tabs";
        DoubleClickCloseMenu.Header = zh ? "\u53cc\u51fb\u5173\u95ed\u6807\u7b7e\u9875" : "Double-click to close tab";
        StartupMenu.Header = zh ? "\u5f00\u673a\u542f\u52a8" : "Start with Windows";
        AutoUpdateMenu.Header = zh ? "\u81ea\u52a8\u66f4\u65b0" : "Automatic updates";
        ShowTrayIconMenu.Header = zh ? "\u663e\u793a\u6258\u76d8\u56fe\u6807" : "Show tray icon";
        CheckUpdatesMenu.Header = zh ? "\u68c0\u67e5\u66f4\u65b0" : "Check for updates";
        ExitMenu.Header = zh ? "\u9000\u51fa" : "Exit";
    }

    public void ApplyThemeText(bool dark, bool zh)
    {
        TrayIcon.ToolTipText = zh
            ? $"WinTab \u5f53\u524d\u4e3a{(dark ? "\u6df1\u8272" : "\u6d45\u8272")}\u754c\u9762\u3002"
            : $"WinTab is using {(dark ? "dark" : "light")} theme.";
    }

    public void RefreshState()
    {
        WindowHookMenu.IsChecked = SettingsManager.IsWindowHookActive;
        ReuseTabsMenu.IsChecked = SettingsManager.ReuseTabs;
        DoubleClickCloseMenu.IsChecked = SettingsManager.DoubleClickCloseTab;
        StartupMenu.IsChecked = RegistryManager.IsStartupEnabled;
        AutoUpdateMenu.IsChecked = SettingsManager.AutoUpdate;
        ShowTrayIconMenu.IsChecked = SettingsManager.ShowTrayIcon;
        ApplyVisibility();
    }

    public void ApplyVisibility()
    {
        TrayIcon.Visibility = SettingsManager.ShowTrayIcon ? Visibility.Visible : Visibility.Collapsed;
    }

    private void WindowHookMenu_Click(object sender, RoutedEventArgs e)
    {
        SettingsManager.IsWindowHookActive = WindowHookMenu.IsChecked;
        _hookManager.SetWindowHook(SettingsManager.IsWindowHookActive);
        RefreshState();
        SettingsChanged?.Invoke(this, EventArgs.Empty);
    }

    private void ReuseTabsMenu_Click(object sender, RoutedEventArgs e)
    {
        SettingsManager.ReuseTabs = ReuseTabsMenu.IsChecked;
        _hookManager.SetReuseTabs(SettingsManager.ReuseTabs);
        RefreshState();
        SettingsChanged?.Invoke(this, EventArgs.Empty);
    }

    private void DoubleClickCloseMenu_Click(object sender, RoutedEventArgs e)
    {
        SettingsManager.DoubleClickCloseTab = DoubleClickCloseMenu.IsChecked;
        _hookManager.SetDoubleClickClose(SettingsManager.DoubleClickCloseTab);
        RefreshState();
        SettingsChanged?.Invoke(this, EventArgs.Empty);
    }

    private void StartupMenu_Click(object sender, RoutedEventArgs e)
    {
        RegistryManager.ToggleStartup();
        if (sender is MenuItem item)
            item.IsChecked = RegistryManager.IsStartupEnabled;

        SettingsChanged?.Invoke(this, EventArgs.Empty);
    }

    private void AutoUpdateMenu_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem item)
            return;

        SettingsManager.AutoUpdate = item.IsChecked;
        SettingsChanged?.Invoke(this, EventArgs.Empty);
    }

    private void ShowTrayIconMenu_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem item)
            return;

        SettingsManager.ShowTrayIcon = item.IsChecked;
        RefreshState();
        SettingsChanged?.Invoke(this, EventArgs.Empty);
    }

    private void OnNotifyIconDoubleClick(object sender, RoutedEventArgs e) => _showWindowAction();

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
        _hookManager.StateChanged -= RefreshState;
        _hookManager.ShellInitialized -= RefreshState;
        TrayIcon.Dispose();
        GC.SuppressFinalize(this);
    }
}
