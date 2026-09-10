using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Navigation;
using System.Windows.Threading;
using WinTab.Helpers;
using WinTab.Managers;
using WinTab.UI.Localization;
using WinTab.UI.Views.Controls;

namespace WinTab.UI.Views;

public partial class MainWindow : Window
{
    // Segoe Fluent Icons / Segoe MDL2 Assets code points for the theme toggle.
    private const string SunGlyph = "\uE706";
    private const string MoonGlyph = "\uE708";

    private readonly HookManager _hookManager;
    private readonly SystemTrayIcon _trayIcon;
    private nint _handle;
    private bool _isExiting;
    private bool _isDisposed;
    private DispatcherTimer? _autoUpdateTimer;
    private DispatcherTimer? _updateFeedbackTimer;
    private Func<string>? _updateFeedback;
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

        MinimizeButton.Click += (_, _) => Hide();
        CloseButton.Click += (_, _) => Hide();
        CheckUpdatesButton.Click += CheckUpdatesButton_Click;
        LanguageToggleButton.Click += LanguageToggleButton_Click;
        ThemeToggleButton.Click += ThemeToggleButton_Click;

        WindowHookToggle.Click += (_, _) => _hookManager.SetWindowHook(WindowHookToggle.IsChecked == true);
        ReuseTabsToggle.Click += (_, _) => _hookManager.SetReuseTabs(ReuseTabsToggle.IsChecked == true);
        DoubleClickCloseToggle.Click += (_, _) => _hookManager.SetDoubleClickClose(DoubleClickCloseToggle.IsChecked == true);
        ShowTrayIconToggle.Click += (_, _) => SettingsManager.ShowTrayIcon = ShowTrayIconToggle.IsChecked == true;
        AutoUpdateToggle.Click += (_, _) => SettingsManager.AutoUpdate = AutoUpdateToggle.IsChecked == true;
        StartupToggle.Click += StartupToggle_Click;

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
        HeroDescriptionText.Text = UiStrings.HeroDescription;
        ExplorerSectionTitleText.Text = UiStrings.ExplorerSectionTitle;
        SystemSectionTitleText.Text = UiStrings.SystemSectionTitle;
        AboutSectionTitleText.Text = UiStrings.AboutSectionTitle;
        BypassHintText.Text = UiStrings.BypassHint;

        WindowHookTitleText.Text = UiStrings.WindowHookTitle;
        WindowHookDescText.Text = UiStrings.WindowHookDescription;
        ReuseTabsTitleText.Text = UiStrings.ReuseTabsTitle;
        ReuseTabsDescText.Text = UiStrings.ReuseTabsDescription;
        DoubleClickTitleText.Text = UiStrings.DoubleClickTitle;
        DoubleClickDescText.Text = UiStrings.DoubleClickDescription;
        StartupTitleText.Text = UiStrings.StartupTitle;
        StartupDescText.Text = UiStrings.StartupDescription;
        ShowTrayIconTitleText.Text = UiStrings.ShowTrayIconTitle;
        ShowTrayIconDescText.Text = SettingsManager.ShowTrayIcon ? UiStrings.ShowTrayIconDescription : UiStrings.ShowTrayIconHiddenDescription;
        AutoUpdateTitleText.Text = UiStrings.AutoUpdateTitle;
        AutoUpdateDescText.Text = UiStrings.AutoUpdateDescription;

        AboutVersionText.Text = $"WinTab v{_appVersion}";
        CheckUpdatesButton.Content = UiStrings.CheckButton;
        ApplyUpdateFeedback();

        LanguageToggleButton.ToolTip = UiStrings.LanguageToggleTooltip;
        ThemeToggleButton.ToolTip = UiStrings.ThemeToggleTooltip(ThemeManager.IsDarkTheme);
        _trayIcon.ApplyLanguage();
    }

    private void ApplyTheme()
    {
        ThemeToggleGlyph.Text = ThemeManager.IsDarkTheme ? SunGlyph : MoonGlyph;
        ThemeToggleButton.ToolTip = UiStrings.ThemeToggleTooltip(ThemeManager.IsDarkTheme);
    }

    private async void CheckUpdatesButton_Click(object sender, RoutedEventArgs e)
    {
        CheckUpdatesButton.IsEnabled = false;
        SetUpdateFeedback(() => UiStrings.UpdateChecking, autoReset: false);

        try
        {
            var result = await UpdateManager.CheckForUpdatesWithResultAsync().ConfigureAwait(true);
            if (!result.Completed)
            {
                SetUpdateFeedback(() => UiStrings.UpdateFailed);
                return;
            }

            if (!result.UpdateAvailable)
            {
                SetUpdateFeedback(() => UiStrings.UpdateUpToDate);
                return;
            }

            if (string.IsNullOrWhiteSpace(result.DownloadUrl))
            {
                SetUpdateFeedback(() => UiStrings.UpdateNoMatchingInstaller);
                return;
            }

            SetUpdateFeedback(() => UiStrings.UpdateOpening);
            UpdateManager.CheckForUpdates();
        }
        finally
        {
            CheckUpdatesButton.IsEnabled = true;
        }
    }

    private void SetUpdateFeedback(Func<string> feedback, bool autoReset = true)
    {
        _updateFeedback = feedback;
        ApplyUpdateFeedback();

        _updateFeedbackTimer?.Stop();
        if (!autoReset)
            return;

        _updateFeedbackTimer ??= new DispatcherTimer(DispatcherPriority.Background, Dispatcher)
        {
            Interval = TimeSpan.FromSeconds(6)
        };
        _updateFeedbackTimer.Tick -= UpdateFeedbackTimer_Tick;
        _updateFeedbackTimer.Tick += UpdateFeedbackTimer_Tick;
        _updateFeedbackTimer.Start();
    }

    private void UpdateFeedbackTimer_Tick(object? sender, EventArgs e)
    {
        _updateFeedbackTimer?.Stop();
        _updateFeedback = null;
        ApplyUpdateFeedback();
    }

    /// <summary>The About row shows the license line until an update check has something to say.</summary>
    private void ApplyUpdateFeedback()
    {
        if (_updateFeedback is { } feedback)
        {
            UpdateFeedbackText.Text = feedback();
            UpdateFeedbackText.Visibility = Visibility.Visible;
            AboutLinksText.Visibility = Visibility.Collapsed;
            return;
        }

        UpdateFeedbackText.Visibility = Visibility.Collapsed;
        AboutLinksText.Visibility = Visibility.Visible;
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
        _updateFeedbackTimer?.Stop();
        Application.Current.Exit -= OnApplicationExit;
        SettingsManager.StaticPropertyChanged -= SettingsManager_StaticPropertyChanged;
        var settingsSaveTask = SettingsManager.FlushSettingsAsync(TimeSpan.FromSeconds(1));
        _trayIcon.Dispose();
        _hookManager.Dispose();
        if (!settingsSaveTask.GetAwaiter().GetResult())
            Trace.TraceError("Could not confirm settings were saved before exit.");
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

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        _handle = new WindowInteropHelper(this).Handle;
    }
}
