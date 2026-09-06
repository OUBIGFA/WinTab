using System;
using System.Windows;
using System.Threading;
using System.Linq;
using System.Windows.Controls;
using WinTab.UI.Views;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.Managers;

namespace WinTab;

// ReSharper disable once RedundantExtendsListEntry
public partial class App : Application
{
    private Mutex? _mutex;
    private EventWaitHandle? _showMainWindowEvent;
    private MainWindow? _mainWindow;
    private bool _isExiting;

    protected override void OnStartup(StartupEventArgs e)
    {
        _mutex = new Mutex(true, Constants.MutexId, out var createdNew);

        if (createdNew)
        {
            base.OnStartup(e);
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            SetupTooltipBehavior();
            ThemeManager.ApplyTheme();
            _showMainWindowEvent = new EventWaitHandle(false, EventResetMode.AutoReset, Constants.ShowMainWindowEventName);

            _mainWindow = new MainWindow();
            StartShowMainWindowRequestListener();

            var launchInBackground = e.Args.Any(arg => string.Equals(arg, Constants.BackgroundLaunchArg, StringComparison.OrdinalIgnoreCase));
            if (!launchInBackground)
                _mainWindow.Show();

            return;
        }

        SignalMainWindow();
        Shutdown();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _isExiting = true;
        _showMainWindowEvent?.Dispose();
        base.OnExit(e);
        if (!ExplorerDebugLog.Complete(TimeSpan.FromMilliseconds(100)))
            System.Diagnostics.Debug.WriteLine("Diagnostic logging is still pending; shutdown will not wait longer.");
        _mutex?.Dispose();
    }

    private void StartShowMainWindowRequestListener()
    {
        var thread = new Thread(() =>
        {
            while (!_isExiting)
            {
                try
                {
                    _showMainWindowEvent?.WaitOne();
                }
                catch (ObjectDisposedException)
                {
                    return;
                }

                if (_isExiting)
                    return;

                Dispatcher.Invoke(() => _mainWindow?.ShowMainWindow());
            }
        })
        {
            IsBackground = true,
            Name = "WinTab show window listener"
        };

        thread.Start();
    }

    private static void SignalMainWindow()
    {
        for (var attempt = 0; attempt < 5; attempt++)
        {
            try
            {
                using var showMainWindowEvent = EventWaitHandle.OpenExisting(Constants.ShowMainWindowEventName);
                showMainWindowEvent.Set();
                return;
            }
            catch (WaitHandleCannotBeOpenedException)
            {
                Thread.Sleep(50);
            }
        }
    }

    private static void SetupTooltipBehavior()
    {
        ToolTipService.ShowDurationProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(3500));
        ToolTipService.InitialShowDelayProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(1700));
        ToolTipService.BetweenShowDelayProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(150));
        ToolTipService.ShowsToolTipOnKeyboardFocusProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(false));
    }
}
