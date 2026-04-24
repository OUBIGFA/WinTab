using System;
using System.Windows;
using System.Threading;
using System.Linq;
using System.Windows.Controls;
using WinTab.UI.Views;
using WinTab.Helpers;
using WinTab.Managers;

namespace WinTab;

// ReSharper disable once RedundantExtendsListEntry
public partial class App : Application
{
    private Mutex? _mutex;
    private MainWindow? _mainWindow;

    protected override void OnStartup(StartupEventArgs e)
    {
        _mutex = new Mutex(true, Constants.MutexId, out var createdNew);

        if (createdNew)
        {
            base.OnStartup(e);
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            SetupTooltipBehavior();
            ThemeManager.ApplyTheme();

            _mainWindow = new MainWindow();
            if (SettingsManager.IsFirstRun)
                SettingsManager.IsFirstRun = false;

            var launchInBackground = e.Args.Any(arg => string.Equals(arg, Constants.BackgroundLaunchArg, StringComparison.OrdinalIgnoreCase));
            if (!launchInBackground)
                _mainWindow.Show();

            return;
        }

        CustomMessageBox.Show("""
                              Another instance is already running.
                              Check in System Tray Icons.
                              """, Constants.AppName, icon: MessageBoxImage.Information);
        Shutdown();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _mutex?.Dispose();
        base.OnExit(e);
    }

    private static void SetupTooltipBehavior()
    {
        ToolTipService.ShowDurationProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(3500));
        ToolTipService.InitialShowDelayProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(1700));
        ToolTipService.BetweenShowDelayProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(150));
        ToolTipService.ShowsToolTipOnKeyboardFocusProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(false));
    }
}
