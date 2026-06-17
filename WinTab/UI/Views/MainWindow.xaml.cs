using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Navigation;
using System.Windows.Media;
using System.Windows.Threading;
using WinTab.Helpers;
using WinTab.Managers;
using WinTab.UI.Views.Controls;
using WinTab.WinAPI;

namespace WinTab.UI.Views;

public partial class MainWindow : Window
{
    private const string LightThemeIconPathData = "M12 18C8.68629 18 6 15.3137 6 12C6 8.68629 8.68629 6 12 6C15.3137 6 18 8.68629 18 12C18 15.3137 15.3137 18 12 18ZM12 16C14.2091 16 16 14.2091 16 12C16 9.79086 14.2091 8 12 8C9.79086 8 8 9.79086 8 12C8 14.2091 9.79086 16 12 16ZM11 1H13V4H11V1ZM11 20H13V23H11V20ZM3.51472 4.92893L4.92893 3.51472L7.05025 5.63604L5.63604 7.05025L3.51472 4.92893ZM16.9497 18.364L18.364 16.9497L20.4853 19.0711L19.0711 20.4853L16.9497 18.364ZM19.0711 3.51472L20.4853 4.92893L18.364 7.05025L16.9497 5.63604L19.0711 3.51472ZM5.63604 16.9497L7.05025 18.364L4.92893 20.4853L3.51472 19.0711L5.63604 16.9497ZM23 11V13H20V11H23ZM4 11V13H1V11H4Z";
    private const string DarkThemeIconPathData = "M10 7C10 10.866 13.134 14 17 14C18.9584 14 20.729 13.1957 21.9995 11.8995C22 11.933 22 11.9665 22 12C22 17.5228 17.5228 22 12 22C6.47715 22 2 17.5228 2 12C2 6.47715 6.47715 2 12 2C12.0335 2 12.067 2 12.1005 2.00049C10.8043 3.27098 10 5.04157 10 7ZM4 12C4 16.4183 7.58172 20 12 20C15.0583 20 17.7158 18.2839 19.062 15.7621C18.3945 15.9187 17.7035 16 17 16C12.0294 16 8 11.9706 8 7C8 6.29648 8.08133 5.60547 8.2379 4.938C5.71611 6.28423 4 8.9417 4 12Z";

    private readonly HookManager _hookManager;
    private readonly SystemTrayIcon _trayIcon;
    private nint _handle;
    private bool _isExiting;
    private bool _isDisposed;
    private bool _isCheckingForUpdates;
    private DispatcherTimer? _autoUpdateTimer;
    private DispatcherTimer? _maintenanceFeedbackTimer;
    private (string Zh, string En)? _maintenanceFeedback;
    private readonly string _appVersion = typeof(MainWindow).Assembly.GetName().Version?.ToString(3) ?? "1.0.0";

    public MainWindow()
    {
        InitializeComponent();

        Width = SettingsManager.FormSize.Width;
        Height = SettingsManager.FormSize.Height;

        _hookManager = new HookManager();
        _trayIcon = new SystemTrayIcon(_hookManager, ShowMainWindow, ExitApplication);

        SetupEventHandlers();
        LoadSettingsIntoUi();
        ApplyThemeText();
        ApplyLanguage();
        _hookManager.ApplySettings();
        RefreshUiState();

        if (SettingsManager.AutoUpdate)
            ScheduleAutomaticUpdateCheck();
    }

    private bool IsChinese => string.Equals(SettingsManager.Language, "zh-CN", StringComparison.OrdinalIgnoreCase);
    private bool IsDarkTheme => ThemeManager.IsDarkTheme;

    private void SetupEventHandlers()
    {
        Application.Current.Exit += OnApplicationExit;
        _hookManager.StateChanged += RefreshUiState;
        _hookManager.ShellInitialized += OnShellInitialized;
        _trayIcon.SettingsChanged += TrayIcon_SettingsChanged;
        SettingsManager.StaticPropertyChanged += SettingsManager_StaticPropertyChanged;

        TitleBar.MouseLeftButtonDown += TitleBar_MouseLeftButtonDown;
        MinimizeButton.Click += (_, _) => HideToTray();
        CloseButton.Click += (_, _) => HideToTray();
        HideWindowButton.Click += (_, _) => HideToTray();
        CheckUpdatesButton.Click += CheckUpdatesButton_Click;
        LanguageToggleButton.Click += LanguageToggleButton_Click;
        ThemeToggleButton.Click += ThemeToggleButton_Click;

        WindowHookToggle.Checked += WindowHookToggle_Changed;
        WindowHookToggle.Unchecked += WindowHookToggle_Changed;
        ReuseTabsToggle.Checked += ReuseTabsToggle_Changed;
        ReuseTabsToggle.Unchecked += ReuseTabsToggle_Changed;
        DoubleClickCloseToggle.Checked += DoubleClickCloseToggle_Changed;
        DoubleClickCloseToggle.Unchecked += DoubleClickCloseToggle_Changed;
        ShowTrayIconToggle.Checked += ShowTrayIconToggle_Changed;
        ShowTrayIconToggle.Unchecked += ShowTrayIconToggle_Changed;
        AutoUpdateToggle.Checked += AutoUpdateToggle_Changed;
        AutoUpdateToggle.Unchecked += AutoUpdateToggle_Changed;
        StartupToggle.Click += StartupToggle_Click;
        CornerResizeThumb.DragDelta += CornerResizeThumb_DragDelta;

        SizeChanged += MainWindow_SizeChanged;
        Closing += MainWindow_Closing;
    }

    private void LoadSettingsIntoUi()
    {
        WindowHookToggle.IsChecked = SettingsManager.IsWindowHookActive;
        ReuseTabsToggle.IsChecked = SettingsManager.ReuseTabs;
        DoubleClickCloseToggle.IsChecked = SettingsManager.DoubleClickCloseTab;
        ShowTrayIconToggle.IsChecked = SettingsManager.ShowTrayIcon;
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

    private void ShowTrayIconToggle_Changed(object sender, RoutedEventArgs e)
    {
        SettingsManager.ShowTrayIcon = ShowTrayIconToggle.IsChecked == true;
        _trayIcon.RefreshState();
        ApplyLanguage();
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
        ShowTrayIconToggle.IsChecked = SettingsManager.ShowTrayIcon;
        _trayIcon.RefreshState();
    }

    private void ScheduleAutomaticUpdateCheck()
    {
        StopAutomaticUpdateCheck();
        _autoUpdateTimer = new DispatcherTimer(DispatcherPriority.ApplicationIdle, Dispatcher)
        {
            Interval = TimeSpan.FromSeconds(2)
        };
        _autoUpdateTimer.Tick += AutomaticUpdateTimer_Tick;
        _autoUpdateTimer.Start();
    }

    private void AutomaticUpdateTimer_Tick(object? sender, EventArgs e)
    {
        StopAutomaticUpdateCheck();
        if (!_isExiting && SettingsManager.AutoUpdate)
            UpdateManager.CheckForUpdates();
    }

    private void StopAutomaticUpdateCheck()
    {
        if (_autoUpdateTimer == null)
            return;

        _autoUpdateTimer.Stop();
        _autoUpdateTimer.Tick -= AutomaticUpdateTimer_Tick;
        _autoUpdateTimer = null;
    }

    private void SettingsManager_StaticPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        SyncSettingsIntoUi();
        ApplyLanguage();
    }

    private void TrayIcon_SettingsChanged(object? sender, EventArgs e)
    {
        SyncSettingsIntoUi();
        ApplyLanguage();
    }

    private void SyncSettingsIntoUi()
    {
        WindowHookToggle.IsChecked = SettingsManager.IsWindowHookActive;
        ReuseTabsToggle.IsChecked = SettingsManager.ReuseTabs;
        DoubleClickCloseToggle.IsChecked = SettingsManager.DoubleClickCloseTab;
        ShowTrayIconToggle.IsChecked = SettingsManager.ShowTrayIcon;
        AutoUpdateToggle.IsChecked = SettingsManager.AutoUpdate;
        StartupToggle.IsChecked = RegistryManager.IsStartupEnabled;
        _trayIcon.ApplyVisibility();
        _trayIcon.RefreshState();
    }

    private void OnShellInitialized()
    {
        RefreshUiState();
    }

    private void ApplyLanguage()
    {
        HeroTitleText.Text = "WinTab";
        HeroDescriptionText.Text = T("\u8ba9\u8d44\u6e90\u7ba1\u7406\u5668\u7a97\u53e3\u56de\u5230\u540c\u4e00\u7ec4\u6807\u7b7e\u3002", "Keep File Explorer windows in one tab set.");
        StatusPillText.Text = T("\u8fd0\u884c\u4e2d", "Running");
        StatusTrayText.Text = SettingsManager.ShowTrayIcon
            ? T("\u53ef\u4ece\u6258\u76d8\u6253\u5f00", "Available from the tray")
            : T("\u6258\u76d8\u56fe\u6807\u5df2\u9690\u85cf", "Tray icon is hidden");
        StatusBypassText.Text = T("\u6309\u4f4f Ctrl + Shift \u53ef\u6253\u5f00\u72ec\u7acb\u7a97\u53e3", "Hold Ctrl + Shift to open a separate window");
        OpenSourceLicenseText.Text = "MIT License";
        OpenSourceVersionText.Text = T($"\u7248\u672c\uff1av{_appVersion}", $"Version: v{_appVersion}");
        OpenSourceLinkText.Text = "GitHub";

        WindowHookTitleText.Text = T("\u5408\u5e76\u65b0\u7a97\u53e3", "Merge new windows");
        WindowHookDescText.Text = T("\u5c06\u65b0\u6253\u5f00\u7684\u6587\u4ef6\u5939\u6536\u56de\u5f53\u524d\u6807\u7b7e\u7ec4\u3002", "Send new folders back to the active Explorer tab group.");
        ReuseTabsTitleText.Text = T("\u590d\u7528\u5df2\u6709\u6807\u7b7e", "Reuse existing tabs");
        ReuseTabsDescText.Text = T("\u8def\u5f84\u5df2\u6253\u5f00\u65f6\uff0c\u76f4\u63a5\u805a\u7126\u5bf9\u5e94\u6807\u7b7e\u3002", "Open a matching path by focusing its current tab.");
        DoubleClickTitleText.Text = T("\u53cc\u51fb\u5173\u95ed\u6807\u7b7e", "Double-click closes tab");
        DoubleClickDescText.Text = T("\u5728\u6807\u9898\u533a\u53cc\u51fb\u5173\u95ed\u5f53\u524d\u6807\u7b7e\u3002", "Close the current Explorer tab from its title area.");
        StartupTitleText.Text = T("\u5f00\u673a\u542f\u52a8", "Start with Windows");
        StartupDescText.Text = T("\u767b\u5f55\u540e\u9759\u9ed8\u8fd0\u884c\u3002", "Run quietly after sign-in.");
        ShowTrayIconTitleText.Text = T("\u663e\u793a\u6258\u76d8\u56fe\u6807", "Show tray icon");
        ShowTrayIconDescText.Text = T("\u5173\u95ed\u7a97\u53e3\u540e\u53ef\u4ece\u901a\u77e5\u533a\u6253\u5f00\u3002", "Keep WinTab available from the notification area.");
        AutoUpdateTitleText.Text = T("\u68c0\u67e5\u66f4\u65b0", "Check for updates");
        AutoUpdateDescText.Text = T("\u6709 GitHub Release \u65f6\u63d0\u793a\u3002", "Notify when a GitHub release is available.");

        ActionsTitleText.Text = T("\u7ef4\u62a4", "Maintenance");
        ApplyMaintenanceDescription();
        CheckUpdatesButton.Content = _isCheckingForUpdates ? T("\u68c0\u67e5\u4e2d", "Checking") : T("\u68c0\u67e5", "Check");
        HideWindowButton.Content = T("\u9690\u85cf", "Hide");

        LanguageToggleButton.ToolTip = IsChinese ? "Switch to English" : "\u5207\u6362\u5230\u4e2d\u6587";
        _trayIcon.ApplyLanguage(IsChinese);
    }

    private void ApplyThemeText()
    {
        ThemeToggleIconPath.Data = Geometry.Parse(IsDarkTheme ? LightThemeIconPathData : DarkThemeIconPathData);
        ThemeToggleButton.ToolTip = IsDarkTheme
            ? T("\u5207\u6362\u4e3a\u6d45\u8272", "Switch to light mode")
            : T("\u5207\u6362\u4e3a\u6df1\u8272", "Switch to dark mode");
        _trayIcon.ApplyThemeText(IsDarkTheme, IsChinese);
    }

    private static string T(string zh, string en)
    {
        return string.Equals(SettingsManager.Language, "zh-CN", StringComparison.OrdinalIgnoreCase) ? zh : en;
    }

    private async void CheckUpdatesButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isCheckingForUpdates)
            return;

        _isCheckingForUpdates = true;
        CheckUpdatesButton.IsEnabled = false;
        CheckUpdatesButton.Content = T("\u68c0\u67e5\u4e2d", "Checking");
        SetMaintenanceFeedback("\u6b63\u5728\u8054\u7f51\u68c0\u67e5\u6700\u65b0\u7248\u672c\u3002", "Checking the latest release online.", autoReset: false);

        try
        {
            var result = await UpdateManager.CheckForUpdatesWithResultAsync().ConfigureAwait(true);
            if (!result.Completed)
            {
                SetMaintenanceFeedback("\u68c0\u67e5\u5931\u8d25\uff0c\u7a0d\u540e\u518d\u8bd5\u3002", "Update check failed. Try again later.");
                return;
            }

            if (!result.UpdateAvailable)
            {
                SetMaintenanceFeedback("\u5f53\u524d\u5df2\u662f\u6700\u65b0\u7248\u672c\u3002", "You're on the latest version.");
                return;
            }

            if (string.IsNullOrWhiteSpace(result.DownloadUrl))
            {
                SetMaintenanceFeedback("\u53d1\u73b0\u65b0\u7248\u672c\uff0c\u4f46\u672a\u627e\u5230\u5339\u914d\u5f53\u524d\u67b6\u6784\u7684\u5b89\u88c5\u5668\u3002", "Update found, but no installer matches this device.");
                return;
            }

            SetMaintenanceFeedback("\u53d1\u73b0\u65b0\u7248\u672c\uff0c\u6b63\u5728\u6253\u5f00\u66f4\u65b0\u7a97\u53e3\u3002", "Update found. Opening the update window.");
            UpdateManager.CheckForUpdates();
        }
        finally
        {
            _isCheckingForUpdates = false;
            CheckUpdatesButton.IsEnabled = true;
            CheckUpdatesButton.Content = T("\u68c0\u67e5", "Check");
        }
    }

    private void SetMaintenanceFeedback(string zh, string en, bool autoReset = true)
    {
        _maintenanceFeedback = (zh, en);
        ApplyMaintenanceDescription();

        _maintenanceFeedbackTimer?.Stop();
        if (!autoReset)
            return;

        _maintenanceFeedbackTimer ??= new DispatcherTimer(DispatcherPriority.Background, Dispatcher)
        {
            Interval = TimeSpan.FromSeconds(6)
        };
        _maintenanceFeedbackTimer.Tick -= MaintenanceFeedbackTimer_Tick;
        _maintenanceFeedbackTimer.Tick += MaintenanceFeedbackTimer_Tick;
        _maintenanceFeedbackTimer.Start();
    }

    private void MaintenanceFeedbackTimer_Tick(object? sender, EventArgs e)
    {
        _maintenanceFeedbackTimer?.Stop();
        _maintenanceFeedback = null;
        ApplyMaintenanceDescription();
    }

    private void ApplyMaintenanceDescription()
    {
        if (_maintenanceFeedback is { } feedback)
        {
            ActionsDescText.Text = T(feedback.Zh, feedback.En);
            return;
        }

        ActionsDescText.Text = SettingsManager.ShowTrayIcon
            ? T("\u5173\u95ed\u6b64\u7a97\u53e3\u540e\uff0cWinTab \u7ee7\u7eed\u7559\u5728\u6258\u76d8\u3002", "The app stays in the tray when this window closes.")
            : T("\u6258\u76d8\u56fe\u6807\u9690\u85cf\u65f6\uff0cWinTab \u4f1a\u76f4\u63a5\u5728\u540e\u53f0\u8fd0\u884c\u3002", "When the tray icon is hidden, WinTab keeps running in the background.");
    }

    private void OpenSourceLink_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
        e.Handled = true;
    }

    public void ShowMainWindow()
    {
        if (WindowState == WindowState.Minimized)
            WindowState = WindowState.Normal;

        Show();
        if (_handle == 0)
            _handle = new WindowInteropHelper(this).Handle;

        Activate();
        Helper.RestoreWindowToForeground(_handle);
    }

    private void HideToTray()
    {
        Hide();
    }

    private void ExitApplication()
    {
        _isExiting = true;
        DisposeApplicationServices();
        Application.Current.Shutdown();
    }

    private void OnApplicationExit(object? sender, ExitEventArgs e)
    {
        DisposeApplicationServices();
    }

    private void DisposeApplicationServices()
    {
        if (_isDisposed)
            return;

        _isDisposed = true;
        StopAutomaticUpdateCheck();
        _maintenanceFeedbackTimer?.Stop();
        Application.Current.Exit -= OnApplicationExit;
        _trayIcon.SettingsChanged -= TrayIcon_SettingsChanged;
        SettingsManager.StaticPropertyChanged -= SettingsManager_StaticPropertyChanged;
        SettingsManager.SaveSettings();
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

    private void CornerResizeThumb_DragDelta(object sender, DragDeltaEventArgs e)
    {
        if (WindowState != WindowState.Normal)
            return;

        Width = Math.Max(MinWidth, Width + e.HorizontalChange);
        Height = Math.Max(MinHeight, Height + e.VerticalChange);
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
