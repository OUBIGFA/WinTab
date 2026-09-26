using System;
using System.Windows;
using System.Threading;
using System.Threading.Tasks;
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
            StartLogging(e.Args);
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

    /// <summary>The log starts before anything else, and failures nobody handles are written to it.</summary>
    private void StartLogging(string[] args)
    {
        var version = typeof(App).Assembly.GetName().Version;
        ExplorerDebugLog.Write($"WinTab {version} started pid={Environment.ProcessId} os={Environment.OSVersion.Version} " +
            $"arch={System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture} args=[{string.Join(" ", args)}]");
        ExplorerDebugLog.Write($"Settings {SettingsManager.Describe()}");
        DispatcherUnhandledException += (_, args) => ExplorerDebugLog.Write($"Unhandled UI exception: {args.Exception}");
        TaskScheduler.UnobservedTaskException += (_, args) => ExplorerDebugLog.Write($"Unobserved task exception: {args.Exception}");
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            ExplorerDebugLog.Write($"Unhandled exception terminating={args.IsTerminating}: {args.ExceptionObject}");
            if (args.IsTerminating)
                ExplorerDebugLog.Complete(TimeSpan.FromMilliseconds(500));
        };
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
