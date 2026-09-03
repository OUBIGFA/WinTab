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
using WinTab.UI.Localization;
using WinTab.UI.Views.Controls;

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
    private Func<string>? _maintenanceFeedback;
    private readonly string _appVersion = typeof(MainWindow).Assembly.GetName().Version?.ToString(3) ?? "1.0.0";

    public MainWindow()
    {
        InitializeComponent();

        Width = SettingsManager.FormSize.Width;
        Height = SettingsManager.FormSize.Height;

        _hookManager = new HookManager();
        _trayIcon = new SystemTrayIcon(_hookManager, ShowMainWindow, ExitApplication);

        SetupEventHandlers();
        SyncSettingsIntoUi();
        ApplyTheme();
        ApplyLanguage();
        _hookManager.ApplySettings();

        if (SettingsManager.AutoUpdate)
            ScheduleAutomaticUpdateCheck();
    }

    private void SetupEventHandlers()
    {
        Application.Current.Exit += OnApplicationExit;
        _hookManager.StateChanged += SyncSettingsIntoUi;
        _hookManager.ShellInitialized += SyncSettingsIntoUi;
        _trayIcon.StartupChanged += (_, _) => SyncSettingsIntoUi();
        SettingsManager.StaticPropertyChanged += SettingsManager_StaticPropertyChanged;

        TitleBar.MouseLeftButtonDown += TitleBar_MouseLeftButtonDown;
        MinimizeButton.Click += (_, _) => Hide();
        CloseButton.Click += (_, _) => Hide();
        HideWindowButton.Click += (_, _) => Hide();
        CheckUpdatesButton.Click += CheckUpdatesButton_Click;
        LanguageToggleButton.Click += LanguageToggleButton_Click;
        ThemeToggleButton.Click += ThemeToggleButton_Click;

        WindowHookToggle.Click += (_, _) => _hookManager.SetWindowHook(WindowHookToggle.IsChecked == true);
        ReuseTabsToggle.Click += (_, _) => _hookManager.SetReuseTabs(ReuseTabsToggle.IsChecked == true);
        DoubleClickCloseToggle.Click += (_, _) => _hookManager.SetDoubleClickClose(DoubleClickCloseToggle.IsChecked == true);
        ShowTrayIconToggle.Click += (_, _) => SettingsManager.ShowTrayIcon = ShowTrayIconToggle.IsChecked == true;
        AutoUpdateToggle.Click += (_, _) => SettingsManager.AutoUpdate = AutoUpdateToggle.IsChecked == true;
        StartupToggle.Click += StartupToggle_Click;
        CornerResizeThumb.DragDelta += CornerResizeThumb_DragDelta;

        SizeChanged += MainWindow_SizeChanged;
        Closing += MainWindow_Closing;
    }

    private void StartupToggle_Click(object sender, RoutedEventArgs e)
    {
        RegistryManager.ToggleStartup();
        SyncSettingsIntoUi();
    }

    private void LanguageToggleButton_Click(object sender, RoutedEventArgs e)
    {
        SettingsManager.Language = UiStrings.IsChinese ? "en-US" : "zh-CN";
    }

    private void ThemeToggleButton_Click(object sender, RoutedEventArgs e)
    {
        SettingsManager.Theme = ThemeManager.IsDarkTheme ? "Light" : "Dark";
        ThemeManager.ApplyTheme();
        ApplyTheme();
    }

    private void SettingsManager_StaticPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        SyncSettingsIntoUi();
        ApplyLanguage();
    }

    /// <summary>Mirrors the persisted settings into every toggle; safe to call from any change source.</summary>
    private void SyncSettingsIntoUi()
    {
        WindowHookToggle.IsChecked = SettingsManager.IsWindowHookActive;
        ReuseTabsToggle.IsChecked = SettingsManager.ReuseTabs;
        DoubleClickCloseToggle.IsChecked = SettingsManager.DoubleClickCloseTab;
        ShowTrayIconToggle.IsChecked = SettingsManager.ShowTrayIcon;
        AutoUpdateToggle.IsChecked = SettingsManager.AutoUpdate;
        StartupToggle.IsChecked = RegistryManager.IsStartupEnabled;
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

    private void ApplyLanguage()
    {
        HeroTitleText.Text = "WinTab";
        HeroDescriptionText.Text = UiStrings.HeroDescription;
        StatusPillText.Text = UiStrings.StatusRunning;
        StatusTrayText.Text = SettingsManager.ShowTrayIcon ? UiStrings.StatusTrayAvailable : UiStrings.StatusTrayHidden;
        StatusBypassText.Text = UiStrings.StatusBypassHint;
        OpenSourceLicenseText.Text = "MIT License";
        OpenSourceVersionText.Text = UiStrings.Version(_appVersion);
        OpenSourceLinkText.Text = "GitHub";

        WindowHookTitleText.Text = UiStrings.WindowHookTitle;
        WindowHookDescText.Text = UiStrings.WindowHookDescription;
        ReuseTabsTitleText.Text = UiStrings.ReuseTabsTitle;
        ReuseTabsDescText.Text = UiStrings.ReuseTabsDescription;
        DoubleClickTitleText.Text = UiStrings.DoubleClickTitle;
        DoubleClickDescText.Text = UiStrings.DoubleClickDescription;
        StartupTitleText.Text = UiStrings.StartupTitle;
        StartupDescText.Text = UiStrings.StartupDescription;
        ShowTrayIconTitleText.Text = UiStrings.ShowTrayIconTitle;
        ShowTrayIconDescText.Text = UiStrings.ShowTrayIconDescription;
        AutoUpdateTitleText.Text = UiStrings.AutoUpdateTitle;
        AutoUpdateDescText.Text = UiStrings.AutoUpdateDescription;

        ActionsTitleText.Text = UiStrings.MaintenanceTitle;
        ApplyMaintenanceDescription();
        CheckUpdatesButton.Content = _isCheckingForUpdates ? UiStrings.CheckingButton : UiStrings.CheckButton;
        HideWindowButton.Content = UiStrings.HideButton;

        LanguageToggleButton.ToolTip = UiStrings.LanguageToggleTooltip;
        ThemeToggleButton.ToolTip = UiStrings.ThemeToggleTooltip(ThemeManager.IsDarkTheme);
        _trayIcon.ApplyLanguage();
    }

    private void ApplyTheme()
    {
        ThemeToggleIconPath.Data = Geometry.Parse(ThemeManager.IsDarkTheme ? LightThemeIconPathData : DarkThemeIconPathData);
        ThemeToggleButton.ToolTip = UiStrings.ThemeToggleTooltip(ThemeManager.IsDarkTheme);
    }

    private async void CheckUpdatesButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isCheckingForUpdates)
            return;

        _isCheckingForUpdates = true;
        CheckUpdatesButton.IsEnabled = false;
        CheckUpdatesButton.Content = UiStrings.CheckingButton;
        SetMaintenanceFeedback(() => UiStrings.UpdateChecking, autoReset: false);

        try
        {
            var result = await UpdateManager.CheckForUpdatesWithResultAsync().ConfigureAwait(true);
            if (!result.Completed)
            {
                SetMaintenanceFeedback(() => UiStrings.UpdateFailed);
                return;
            }

            if (!result.UpdateAvailable)
            {
                SetMaintenanceFeedback(() => UiStrings.UpdateUpToDate);
                return;
            }

            if (string.IsNullOrWhiteSpace(result.DownloadUrl))
            {
                SetMaintenanceFeedback(() => UiStrings.UpdateNoMatchingInstaller);
                return;
            }

            SetMaintenanceFeedback(() => UiStrings.UpdateOpening);
            UpdateManager.CheckForUpdates();
        }
        finally
        {
            _isCheckingForUpdates = false;
            CheckUpdatesButton.IsEnabled = true;
            CheckUpdatesButton.Content = UiStrings.CheckButton;
        }
    }

    private void SetMaintenanceFeedback(Func<string> feedback, bool autoReset = true)
    {
        _maintenanceFeedback = feedback;
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
        if (_maintenanceFeedback != null)
        {
            ActionsDescText.Text = _maintenanceFeedback();
            return;
        }

        ActionsDescText.Text = SettingsManager.ShowTrayIcon ? UiStrings.MaintenanceTrayVisible : UiStrings.MaintenanceTrayHidden;
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
        Hide();
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
