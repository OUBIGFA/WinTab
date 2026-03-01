using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using WinTab.Diagnostics;
using WinTab.Platform.Win32;

namespace WinTab.App.Services;

public sealed class SingleInstanceService : IDisposable
{
    private const string ActivationEventName = "WinTab_ActivateMainWindow";
    private readonly string _mutexName;
    private readonly Logger? _logger;
    private Mutex? _singleInstanceMutex;
    private EventWaitHandle? _activationEvent;
    private CancellationTokenSource? _activationListenerCts;
    private Task? _activationListenerTask;

    public bool OwnsMutex { get; private set; }

    public SingleInstanceService(string mutexName = "WinTab_SingleInstance", Logger? logger = null)
    {
        _mutexName = mutexName;
        _logger = logger;
    }

    /// <summary>
    /// Attempts to acquire the single instance mutex.
    /// Returns true if this is the first instance.
    /// </summary>
    public bool InitializeAsFirstInstance()
    {
        _singleInstanceMutex = new Mutex(true, _mutexName, out bool isNewInstance);
        OwnsMutex = isNewInstance;
        return isNewInstance;
    }

    /// <summary>
    /// Starts a background listener waiting for activation signals from other instances.
    /// </summary>
    public void StartActivationListener()
    {
        _activationEvent = new EventWaitHandle(false, EventResetMode.AutoReset, ActivationEventName);
        _activationListenerCts = new CancellationTokenSource();
        var token = _activationListenerCts.Token;
        var activationEvent = _activationEvent;

        _activationListenerTask = Task.Run(() =>
        {
            while (!token.IsCancellationRequested)
            {
                bool signaled = activationEvent.WaitOne(millisecondsTimeout: 200);
                if (token.IsCancellationRequested)
                    break;

                if (!signaled)
                    continue;

                try
                {
                    System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                    {
                        if (System.Windows.Application.Current?.MainWindow is not Window window)
                            return;

                        if (!window.IsVisible)
                            window.Show();

                        window.ShowInTaskbar = true;
                        window.WindowState = WindowState.Normal;
                        window.Activate();
                    });
                }
                catch (InvalidOperationException ex)
                {
                    _logger?.Error("Failed to activate main window from another instance signal.", ex);
                }
            }
        }, token);
    }

    /// <summary>
    /// Signals an already running instance to activate its main window.
    /// </summary>
    public void SignalExistingInstanceActivation()
    {
        try
        {
            using var activationEvent = EventWaitHandle.OpenExisting(ActivationEventName);
            activationEvent.Set();
        }
        catch (WaitHandleCannotBeOpenedException ex)
        {
            _logger?.Warn($"Could not signal existing instance (it might not be listening yet). {ex.Message}");
        }
        catch (UnauthorizedAccessException ex)
        {
            _logger?.Warn($"Could not signal existing instance (Access denied). {ex.Message}");
        }
    }

    /// <summary>
    /// Attempts to bring the existing instance's process window to the foreground.
    /// </summary>
    public void BringExistingInstanceToForeground()
    {
        try
        {
            using var current = Process.GetCurrentProcess();
            var existing = Process.GetProcessesByName(current.ProcessName)
                .Where(p => p.Id != current.Id)
                .OrderBy(p => p.StartTime)
                .FirstOrDefault(p => p.MainWindowHandle != IntPtr.Zero);

            if (existing is null)
                return;

            IntPtr hWnd = existing.MainWindowHandle;
            if (hWnd == IntPtr.Zero)
                return;

            NativeMethods.ShowWindow(hWnd, NativeConstants.SW_RESTORE);
            NativeMethods.SetForegroundWindow(hWnd);
        }
        catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception)
        {
            _logger?.Warn($"Failed to bring existing instance to foreground. {ex.Message}");
        }
    }

    public void Dispose()
    {
        if (_activationListenerCts is not null)
        {
            _activationListenerCts.Cancel();
        }

        if (_activationListenerTask is not null)
        {
            try
            {
                _activationListenerTask.Wait(500);
            }
            catch (AggregateException ex) when (ex.InnerExceptions.All(e => e is TaskCanceledException or OperationCanceledException))
            {
                // ignore
            }
            catch (OperationCanceledException)
            {
                // ignore
            }
            finally
            {
                _activationListenerTask = null;
            }
        }

        if (_activationListenerCts is not null)
        {
            _activationListenerCts.Dispose();
            _activationListenerCts = null;
        }

        if (_activationEvent is not null)
        {
            _activationEvent.Dispose();
            _activationEvent = null;
        }

        if (OwnsMutex && _singleInstanceMutex is not null)
        {
            try { _singleInstanceMutex.ReleaseMutex(); }
            catch { /* ignore */ }
        }

        if (_singleInstanceMutex is not null)
        {
            _singleInstanceMutex.Dispose();
            _singleInstanceMutex = null;
        }
    }
}
