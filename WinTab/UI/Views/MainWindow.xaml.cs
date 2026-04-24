using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Navigation;
using WinTab.Managers;
using WinTab.UI.Views.Controls;
using WinTab.WinAPI;

namespace WinTab.UI.Views;

public partial class MainWindow : Window
{
    private readonly HookManager _hookManager;
    private readonly SystemTrayIcon _trayIcon;
    private nint _handle;
    private bool _isExiting;

    public MainWindow()
    {
        InitializeComponent();

        Width = SettingsManager.FormSize.Width;
        Height = SettingsManager.FormSize.Height;

        _hookManager = new HookManager();
        _trayIcon = new SystemTrayIcon(_hookManager, ShowFromTray, ExitApplication);

        SetupEventHandlers();
        LoadSettingsIntoUi();
        ApplyThemeText();
        ApplyLanguage();
        _hookManager.ApplySettings();
        RefreshUiState();

        if (SettingsManager.AutoUpdate)
            UpdateManager.CheckForUpdates();
    }

    private bool IsChinese => string.Equals(SettingsManager.Language, "zh-CN", StringComparison.OrdinalIgnoreCase);
    private bool IsDarkTheme => ThemeManager.IsDarkTheme;

    private void SetupEventHandlers()
    {
        Application.Current.Exit += OnApplicationExit;
        _hookManager.StateChanged += RefreshUiState;
        _hookManager.ShellInitialized += OnShellInitialized;

        TitleBar.MouseLeftButtonDown += TitleBar_MouseLeftButtonDown;
        MinimizeButton.Click += (_, _) => HideToTray();
        CloseButton.Click += (_, _) => HideToTray();
        HideWindowButton.Click += (_, _) => HideToTray();
        CheckUpdatesButton.Click += (_, _) => UpdateManager.CheckForUpdates();
        LanguageToggleButton.Click += LanguageToggleButton_Click;
        ThemeToggleButton.Click += ThemeToggleButton_Click;

        WindowHookToggle.Checked += WindowHookToggle_Changed;
        WindowHookToggle.Unchecked += WindowHookToggle_Changed;
        ReuseTabsToggle.Checked += ReuseTabsToggle_Changed;
        ReuseTabsToggle.Unchecked += ReuseTabsToggle_Changed;
        DoubleClickCloseToggle.Checked += DoubleClickCloseToggle_Changed;
        DoubleClickCloseToggle.Unchecked += DoubleClickCloseToggle_Changed;
        AutoUpdateToggle.Checked += AutoUpdateToggle_Changed;
        AutoUpdateToggle.Unchecked += AutoUpdateToggle_Changed;
        StartupToggle.Click += StartupToggle_Click;

        SizeChanged += MainWindow_SizeChanged;
        Closing += MainWindow_Closing;
    }

    private void LoadSettingsIntoUi()
    {
        WindowHookToggle.IsChecked = SettingsManager.IsWindowHookActive;
        ReuseTabsToggle.IsChecked = SettingsManager.ReuseTabs;
        DoubleClickCloseToggle.IsChecked = SettingsManager.DoubleClickCloseTab;
        AutoUpdateToggle.IsChecked = SettingsManager.AutoUpdate;
        StartupToggle.IsChecked = RegistryManager.IsStartupEnabled;
    }

    private void WindowHookToggle_Changed(object sender, RoutedEventArgs e)
    {
        SettingsManager.IsWindowHookActive = WindowHookToggle.IsChecked == true;
        _hookManager.SetWindowHook(SettingsManager.IsWindowHookActive);
        ReuseTabsToggle.IsChecked = SettingsManager.ReuseTabs;
    }

    private void ReuseTabsToggle_Changed(object sender, RoutedEventArgs e)
    {
        SettingsManager.ReuseTabs = ReuseTabsToggle.IsChecked == true;
        _hookManager.SetReuseTabs(SettingsManager.ReuseTabs);
        WindowHookToggle.IsChecked = SettingsManager.IsWindowHookActive;
    }

    private void DoubleClickCloseToggle_Changed(object sender, RoutedEventArgs e)
    {
        SettingsManager.DoubleClickCloseTab = DoubleClickCloseToggle.IsChecked == true;
        _hookManager.SetDoubleClickClose(SettingsManager.DoubleClickCloseTab);
    }

    private void AutoUpdateToggle_Changed(object sender, RoutedEventArgs e)
    {
        SettingsManager.AutoUpdate = AutoUpdateToggle.IsChecked == true;
    }

    private void StartupToggle_Click(object sender, RoutedEventArgs e)
    {
        RegistryManager.ToggleStartup();
        StartupToggle.IsChecked = RegistryManager.IsStartupEnabled;
    }

    private void LanguageToggleButton_Click(object sender, RoutedEventArgs e)
    {
        SettingsManager.Language = IsChinese ? "en-US" : "zh-CN";
        ApplyThemeText();
        ApplyLanguage();
        RefreshUiState();
    }

    private void ThemeToggleButton_Click(object sender, RoutedEventArgs e)
    {
        SettingsManager.Theme = IsDarkTheme ? "Light" : "Dark";
        ThemeManager.ApplyTheme();
        ApplyThemeText();
    }

    private void RefreshUiState()
    {
        StartupToggle.IsChecked = RegistryManager.IsStartupEnabled;
        _trayIcon.RefreshState();
    }

    private void OnShellInitialized()
    {
        RefreshUiState();
    }

    private void ApplyLanguage()
    {
        HeroTitleText.Text = T("\u8d44\u6e90\u7ba1\u7406\u5668\u6807\u7b7e\u7ba1\u7406\u5de5\u5177", "File Explorer tab manager");
        HeroDescriptionText.Text = T("WinTab \u662f\u4e00\u4e2a\u4e13\u6ce8\u4e8e Windows 11 \u8d44\u6e90\u7ba1\u7406\u5668\u6807\u7b7e\u9875\u4f53\u9a8c\u7684\u5c0f\u5de5\u5177\u3002\u5b83\u4f1a\u628a\u65b0\u6253\u5f00\u7684\u6587\u4ef6\u5939\u4f18\u5148\u5e76\u5165\u73b0\u6709\u8d44\u6e90\u7ba1\u7406\u5668\u7a97\u53e3\uff0c\u5e76\u5728\u76ee\u6807\u8def\u5f84\u5df2\u7ecf\u6253\u5f00\u65f6\u590d\u7528\u5df2\u6709\u6807\u7b7e\u9875\uff0c\u51cf\u5c11\u91cd\u590d\u7a97\u53e3\u548c\u91cd\u590d\u6807\u7b7e\u3002", "WinTab is a focused utility for File Explorer tabs on Windows 11. It merges newly opened folders into existing Explorer windows and reuses tabs when the target path is already open.");
        AppTaglineText.Text = T("\u8d44\u6e90\u7ba1\u7406\u5668\u6807\u7b7e\u63a7\u5236", "Explorer tab control");
        OpenSourceLicenseText.Text = T("\u5f00\u6e90\u534f\u8bae\uff1aMIT License", "Open source license: MIT License");
        OpenSourceLinkText.Text = "https://github.com/OUBIGFA/WinTab";
        CoreSwitchesTitleText.Text = T("\u6838\u5fc3\u5f00\u5173", "Core switches");

        WindowHookTitleText.Text = T("\u81ea\u52a8\u5408\u5e76\u65b0\u7684\u8d44\u6e90\u7ba1\u7406\u5668\u7a97\u53e3", "Auto-merge new Explorer windows");
        WindowHookDescText.Text = T("\u65b0\u5f00\u7684\u6587\u4ef6\u5939\u4f1a\u4f18\u5148\u5e76\u5165\u5f53\u524d\u8d44\u6e90\u7ba1\u7406\u5668\u4f1a\u8bdd\u3002", "Keep new folder opens inside the existing Explorer session.");
        ReuseTabsTitleText.Text = T("\u590d\u7528\u540c\u8def\u5f84\u6807\u7b7e\u9875", "Reuse matching tabs");
        ReuseTabsDescText.Text = T("\u4f18\u5148\u5207\u6362\u5230\u5df2\u6709\u540c\u8def\u5f84\u6807\u7b7e\u9875\uff0c\u907f\u514d\u91cd\u590d\u6253\u5f00\u3002", "Switch to an existing folder tab instead of duplicating it.");
        DoubleClickTitleText.Text = T("\u53cc\u51fb\u6807\u7b7e\u9875\u6807\u9898\u5173\u95ed\u6807\u7b7e\u9875", "Double-click tab title to close");
        DoubleClickDescText.Text = T("\u4ec5\u5728 Explorer \u6807\u7b7e\u6807\u9898\u533a\u57df\u751f\u6548\uff0c\u7a7a\u767d\u6807\u9898\u680f\u4fdd\u6301\u7cfb\u7edf\u539f\u751f\u884c\u4e3a\u3002", "Only acts on Explorer tab headers. Blank title bar keeps native behavior.");
        StartupTitleText.Text = T("\u5f00\u673a\u542f\u52a8", "Start with Windows");
        StartupDescText.Text = T("WinTab \u4f1a\u5728\u5f53\u524d Windows \u7528\u6237\u767b\u5f55\u540e\u542f\u52a8\u3002", "Starts WinTab after sign-in for the current Windows user.");
        AutoUpdateTitleText.Text = T("\u81ea\u52a8\u66f4\u65b0", "Automatic updates");
        AutoUpdateDescText.Text = T("\u542f\u52a8\u65f6\u68c0\u67e5 WinTab \u65b0\u7248\u672c\u3002", "Checks WinTab releases on startup.");

        ActionsTitleText.Text = T("\u64cd\u4f5c", "Actions");
        ActionsDescText.Text = T("\u53ef\u968f\u65f6\u4ece\u6258\u76d8\u6253\u5f00\u3002\u5173\u95ed\u7a97\u53e3\u4e0d\u4f1a\u9000\u51fa WinTab\u3002", "Open from tray anytime. Closing this window keeps WinTab running.");
        CheckUpdatesButton.Content = T("\u68c0\u67e5\u66f4\u65b0", "Check Updates");
        HideWindowButton.Content = T("\u9690\u85cf\u5230\u6258\u76d8", "Hide to Tray");
        FooterTipText.Text = T("\u6253\u5f00\u6587\u4ef6\u5939\u65f6\u6309\u4f4f Ctrl + Shift\uff0c\u53ef\u4fdd\u6301\u4e3a\u72ec\u7acb\u8d44\u6e90\u7ba1\u7406\u5668\u7a97\u53e3\u3002", "Ctrl + Shift while opening a folder keeps that open as a separate Explorer window.");

        LanguageToggleButton.ToolTip = IsChinese ? "Switch to English" : "\u5207\u6362\u5230\u4e2d\u6587";
        LanguageToggleIcon.Text = "🌐";
        _trayIcon.ApplyLanguage(IsChinese);
    }

    private void ApplyThemeText()
    {
        ThemeToggleIcon.Text = IsDarkTheme ? "☀" : "🌙";
        ThemeToggleButton.ToolTip = IsDarkTheme
            ? T("\u5207\u6362\u4e3a\u6d45\u8272", "Switch to light mode")
            : T("\u5207\u6362\u4e3a\u6df1\u8272", "Switch to dark mode");
        _trayIcon.ApplyThemeText(IsDarkTheme, IsChinese);
    }

    private static string T(string zh, string en)
    {
        return string.Equals(SettingsManager.Language, "zh-CN", StringComparison.OrdinalIgnoreCase) ? zh : en;
    }

    private void OpenSourceLink_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
        e.Handled = true;
    }

    private void ShowFromTray()
    {
        Show();
        if (_handle == 0)
            _handle = new WindowInteropHelper(this).Handle;

        Activate();
        WinApi.RestoreWindowToForeground(_handle);
    }

    private void HideToTray()
    {
        Hide();
    }

    private void ExitApplication()
    {
        _isExiting = true;
        Application.Current.Exit -= OnApplicationExit;
        _trayIcon.Dispose();
        _hookManager.Dispose();
        Application.Current.Shutdown();
    }

    private void OnApplicationExit(object? sender, ExitEventArgs e)
    {
        _trayIcon.Dispose();
        _hookManager.Dispose();
    }

    private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (WindowState == WindowState.Normal)
            SettingsManager.FormSize = new Size(Width, Height);
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_isExiting)
            return;

        e.Cancel = true;
        HideToTray();
    }

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
        {
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
            return;
        }

        DragMove();
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        _handle = new WindowInteropHelper(this).Handle;
    }
}
