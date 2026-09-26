using System;
using System.Windows;
using System.Windows.Controls;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.Managers;
using WinTab.UI.Localization;

namespace WinTab.UI.Views.Controls;

public partial class SystemTrayIcon : UserControl, IDisposable
{
    private readonly HookManager _hookManager;
    private readonly Action _showWindowAction;
    private readonly Action _exitAction;
    private bool _disposed;

    /// <summary>
    /// Startup lives in the registry, not in <see cref="SettingsManager"/>, so it is the only
    /// tray change the main window cannot learn about from <see cref="SettingsManager.StaticPropertyChanged"/>.
    /// </summary>
    public event EventHandler? StartupChanged;

    public SystemTrayIcon(HookManager hookManager, Action showWindowAction, Action exitAction)
    {
        InitializeComponent();

        _hookManager = hookManager;
        _showWindowAction = showWindowAction;
        _exitAction = exitAction;

        TrayIcon.Icon = Helper.GetIcon();
        ApplyLanguage();

        OpenSettingsMenu.Click += (_, _) => _showWindowAction();
        WindowHookMenu.Click += (_, _) => _hookManager.SetWindowHook(WindowHookMenu.IsChecked);
        ReuseTabsMenu.Click += (_, _) => _hookManager.SetReuseTabs(ReuseTabsMenu.IsChecked);
        DoubleClickCloseMenu.Click += (_, _) => _hookManager.SetDoubleClickClose(DoubleClickCloseMenu.IsChecked);
        MiddleClickForegroundMenu.Click += (_, _) => _hookManager.SetMiddleClickForeground(MiddleClickForegroundMenu.IsChecked);
        WheelSwitchMenu.Click += (_, _) => _hookManager.SetWheelSwitch(WheelSwitchMenu.IsChecked);
        RestoreGroupMenu.Click += async (_, _) => await _hookManager.ExecuteSessionCommandAsync(true);
        ReopenTabMenu.Click += async (_, _) => await _hookManager.ExecuteSessionCommandAsync(false);
        RecordClosedTabsMenu.Click += (_, _) => _hookManager.SetReopenClosedTab(RecordClosedTabsMenu.IsChecked);
        _hookManager.SessionCommandFinished += SessionCommandFinished;
        StartupMenu.Click += StartupMenu_Click;
        AutoUpdateMenu.Click += (_, _) => SettingsManager.AutoUpdate = AutoUpdateMenu.IsChecked;
        ShowTrayIconMenu.Click += (_, _) => SettingsManager.ShowTrayIcon = ShowTrayIconMenu.IsChecked;
        CheckUpdatesMenu.Click += (_, _) => UpdateManager.CheckForUpdates();
        ExitMenu.Click += (_, _) => _exitAction();

        _hookManager.StateChanged += RefreshState;
        _hookManager.ShellInitialized += RefreshState;
        RefreshState();
    }

    public void ApplyLanguage()
    {
        TrayIcon.ToolTipText = UiStrings.TrayTooltip;
        OpenSettingsMenu.Header = UiStrings.TrayOpen;
        WindowHookMenu.Header = UiStrings.TrayWindowHook;
        ReuseTabsMenu.Header = UiStrings.TrayReuseTabs;
        DoubleClickCloseMenu.Header = UiStrings.TrayDoubleClickClose;
        MiddleClickForegroundMenu.Header = UiStrings.TrayMiddleClickForeground;
        WheelSwitchMenu.Header = UiStrings.TrayWheelSwitch;
        RestoreGroupMenu.Header = UiStrings.RestoreGroupCommand;
        ReopenTabMenu.Header = UiStrings.ReopenTabCommand;
        RecordClosedTabsMenu.Header = UiStrings.RecordClosedTabs;
        StartupMenu.Header = UiStrings.TrayStartup;
        AutoUpdateMenu.Header = UiStrings.TrayAutoUpdate;
        ShowTrayIconMenu.Header = UiStrings.TrayShowTrayIcon;
        CheckUpdatesMenu.Header = UiStrings.TrayCheckUpdates;
        ExitMenu.Header = UiStrings.TrayExit;
    }

    public void RefreshState()
    {
        WindowHookMenu.IsChecked = SettingsManager.IsWindowHookActive;
        ReuseTabsMenu.IsChecked = SettingsManager.ReuseTabs;
        DoubleClickCloseMenu.IsChecked = SettingsManager.DoubleClickCloseTab;
        MiddleClickForegroundMenu.IsChecked = SettingsManager.MiddleClickForegroundTab;
        WheelSwitchMenu.IsChecked = SettingsManager.WheelSwitchTab;
        RestoreGroupMenu.IsEnabled = _hookManager.IsShellReady;
        ReopenTabMenu.IsEnabled = _hookManager.IsShellReady && SettingsManager.ReopenClosedTab;
        RecordClosedTabsMenu.IsChecked = SettingsManager.ReopenClosedTab;
        RestoreGroupMenu.InputGestureText = SettingsManager.RestoreGroupShortcutEnabled ? SettingsManager.RestoreGroupShortcut : string.Empty;
        ReopenTabMenu.InputGestureText = SettingsManager.ReopenTabShortcutEnabled && SettingsManager.ReopenClosedTab ? SettingsManager.ReopenTabShortcut : string.Empty;
        StartupMenu.IsChecked = RegistryManager.IsStartupEnabled;
        AutoUpdateMenu.IsChecked = SettingsManager.AutoUpdate;
        ShowTrayIconMenu.IsChecked = SettingsManager.ShowTrayIcon;
        TrayIcon.Visibility = SettingsManager.ShowTrayIcon ? Visibility.Visible : Visibility.Collapsed;
    }

    private void SessionCommandFinished(SessionCommandResult result)
    {
        // Successful recovery is already visible in Explorer; avoid stealing focus with a success dialog.
        if (_disposed || result == SessionCommandResult.Completed) return;
        if (SettingsManager.ShowTrayIcon)
            TrayIcon.ShowBalloonTip(UiStrings.RecoveryTitle, UiStrings.SessionResult(result), Hardcodet.Wpf.TaskbarNotification.BalloonIcon.Info);
        else
            MessageBox.Show(UiStrings.SessionResult(result), UiStrings.RecoveryTitle, MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void StartupMenu_Click(object sender, RoutedEventArgs e)
    {
        RegistryManager.ToggleStartup();
        StartupMenu.IsChecked = RegistryManager.IsStartupEnabled;
        StartupChanged?.Invoke(this, EventArgs.Empty);
    }

    private void OnNotifyIconDoubleClick(object sender, RoutedEventArgs e) => _showWindowAction();

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
        _hookManager.SessionCommandFinished -= SessionCommandFinished;
        _hookManager.StateChanged -= RefreshState;
        _hookManager.ShellInitialized -= RefreshState;
        TrayIcon.Dispose();
        GC.SuppressFinalize(this);
    }
}
