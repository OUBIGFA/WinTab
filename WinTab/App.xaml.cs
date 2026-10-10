using System;
using System.Windows;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using System.Windows.Controls;
using System.Runtime.InteropServices;
using WinTab.UI.Desktop;
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
    private EventWaitHandle? _openRecycleBinEvent;
    private IApplicationUi? _applicationUi;
    private bool _isExiting;

    protected override void OnStartup(StartupEventArgs e)
    {
        var openRecycleBin = e.Args.Any(arg => string.Equals(arg, Constants.OpenRecycleBinArg, StringComparison.OrdinalIgnoreCase));
        var launchInBackground = openRecycleBin || e.Args.Any(arg => string.Equals(arg, Constants.BackgroundLaunchArg, StringComparison.OrdinalIgnoreCase));
        _mutex = new Mutex(true, Constants.MutexId, out var createdNew);

        if (createdNew)
        {
            StartLogging(e.Args);
            base.OnStartup(e);
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            SetupTooltipBehavior();
            ThemeManager.ApplyTheme();
            _showMainWindowEvent = new EventWaitHandle(false, EventResetMode.AutoReset, Constants.ShowMainWindowEventName);
            _openRecycleBinEvent = new EventWaitHandle(false, EventResetMode.AutoReset, Constants.OpenRecycleBinEventName);

            // MyGo currently targets Windows x64/arm64; x86 keeps the existing WPF interface.
            // Do not silently substitute WPF when a 64-bit installation is missing its UI payload.
            _applicationUi = RuntimeInformation.ProcessArchitecture == Architecture.X86
                ? new MainWindow()
                : new DesktopApplication();
            StartShowMainWindowRequestListener();

            if (!launchInBackground)
                _applicationUi.ShowMainWindow();
            if (openRecycleBin) _openRecycleBinEvent.Set();

            return;
        }

        // A background restart must also stay quiet when another instance is already running.
        if (launchInBackground && !openRecycleBin)
        {
            Shutdown();
            return;
        }

        // The process launched by Shell owns the foreground grant; pass it on before signalling the resident app.
        AllowSetForegroundWindow(uint.MaxValue);
        if (!SignalRequest(openRecycleBin ? Constants.OpenRecycleBinEventName : Constants.ShowMainWindowEventName))
        {
            System.Diagnostics.Debug.WriteLine("The resident WinTab instance did not accept the launch request");
            if (openRecycleBin) HookManager.OpenRecycleBinNatively();
        }
        Shutdown();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _isExiting = true;
        _showMainWindowEvent?.Set();
        _showMainWindowEvent?.Dispose();
        _openRecycleBinEvent?.Dispose();
        base.OnExit(e);
        ExplorerDebugLog.Write($"WinTab exiting code={e.ApplicationExitCode} pendingWindowRecovery={ExplorerWindowVisibility.HiddenWindowHandles.Count()}");
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
            var requests = new WaitHandle[] { _showMainWindowEvent!, _openRecycleBinEvent! };
            while (!_isExiting)
            {
                int request;
                try
                {
                    request = WaitHandle.WaitAny(requests);
                }
                catch (ObjectDisposedException)
                {
                    return;
                }

                if (_isExiting)
                    return;

                Dispatcher.InvokeAsync(async () =>
                {
                    if (_isExiting || _applicationUi == null) return;
                    if (request == 0) _applicationUi.ShowMainWindow();
                    else await _applicationUi.OpenRecycleBinAsync();
                });
            }
        })
        {
            IsBackground = true,
            Name = "WinTab show window listener"
        };

        thread.Start();
    }

    private static bool SignalRequest(string eventName)
    {
        for (var attempt = 0; attempt < 5; attempt++)
        {
            try
            {
                using var showMainWindowEvent = EventWaitHandle.OpenExisting(eventName);
                showMainWindowEvent.Set();
                return true;
            }
            catch (WaitHandleCannotBeOpenedException)
            {
                Thread.Sleep(50);
            }
        }
        return false;
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool AllowSetForegroundWindow(uint processId);

    private static void SetupTooltipBehavior()
    {
        ToolTipService.ShowDurationProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(3500));
        ToolTipService.InitialShowDelayProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(1700));
        ToolTipService.BetweenShowDelayProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(150));
        ToolTipService.ShowsToolTipOnKeyboardFocusProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(false));
    }
}
