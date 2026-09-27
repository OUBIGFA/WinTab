using SHDocVw;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Interop;
using WinTab.Models;

namespace WinTab.Hooks;

using WindowEntry = DualKeyEntry<InternetExplorer, nint?, WindowInfo>;

/// <summary>
/// The connection to the Explorer shell: finding its process, creating and releasing the shell objects,
/// and running shell calls on the STA worker.
/// </summary>
public partial class ExplorerWatcher
{
    private Task RunInStaThread(Action action, TaskCreationOptions options = default, CancellationToken cancellationToken = default)
    {
        return RunInStaThread(() =>
        {
            action();
            return true;
        }, options, cancellationToken);
    }

    private Task<T?> RunInStaThread<T>(Func<T?> action, TaskCreationOptions options = default, CancellationToken cancellationToken = default)
    {
        var generation = _shellGeneration;
        var token = cancellationToken.CanBeCanceled ? cancellationToken : CurrentCancellation;
        T? Execute()
        {
            token.ThrowIfCancellationRequested();
            EnsureCurrentMerge();
            if (generation != _shellGeneration)
                throw new OperationCanceledException("The Explorer connection has changed.");
            return action();
        }
        return _staTaskScheduler.IsCurrentThread
            ? Task.FromResult(Execute())
            : Task.Factory.StartNew(Execute, token, options, _staTaskScheduler);
    }

    private void StartExplorerProcessCheck() => _explorerCheckTimer = new Timer(CheckForMainExplorer, null, 0, 1000);

    private void CheckForMainExplorer(object? state)
    {
        if (_disposed)
            return;
        if (_mainExplorerProcessId != 0)
        {
            if (HasPendingTabRegistrations() || HasUnregisteredExplorerTabs())
                ScheduleShellWindowRegistration();
            return;
        }
        if (Interlocked.CompareExchange(ref _shellTransitionScheduled, 1, 0) != 0)
            return;
        using var process = ExplorerWindowDiscovery.GetMainExplorerProcess();
        if (process == null)
        {
            Volatile.Write(ref _shellTransitionScheduled, 0);
            return;
        }
        _shellTransitionTask = InitializeShellAsync(process.Id);
    }

    private void OnExplorerProcessTerminated(object? sender, ProcessEventArgs eventArgs)
    {
        if (_disposed)
            return;
        if (eventArgs.ProcessId == _mainExplorerProcessId)
        {
            RetireShellConnection("explorer-restarted");
        }
        else
        {
            ScheduleShellWindowRegistration();
        }
    }

    private async Task InitializeShellObjectsAsync()
    {
        _preExistingExplorerWindowsProtected = false;
        _shellPathComparer = new ShellPathComparer();
        _shellWindows = new ShellWindows();
        _shellLifetime.Token.ThrowIfCancellationRequested();

        _defaultLocation = GetDefaultExplorerLocation();
        ClearShellCaches();
        if (_captureSessions)
        {
            // Pure snapshots survive a transient COM disconnect; their HWND/process identities decide
            // whether they closed normally or belong to an ended Explorer on the next capture pass.
            _sessionVisuals?.Clear();
            ExcludeExistingWindowsFromRestore();
        }
        RecoverHiddenExplorerWindows("initialize-shell");

        if (ExplorerWindowDiscovery.IsFileExplorerForeground(out var foregroundWindow))
            MainWindowHandle = foregroundWindow;

        // Hook the global "WindowRegistered" event
        _windowRegisteredHandler = OnShellWindowRegistered;
        _shellWindows.WindowRegistered += _windowRegisteredHandler;

        // WinEvent only wakes ShellWindows processing; WindowRegistered owns merge and release.
        _shellLifetime.Token.ThrowIfCancellationRequested();
        _eventObjectShowHookCallback = OnWindowShown;
        _winEventHookThread = new WinEventHookThread(_eventObjectShowHookCallback);
        _winEventHookThread.Start();

        // Hook the event handlers for already-open windows
        var registrations = new List<Task>();
        var count = _shellWindows.Count;
        for (var i = 0; i < count; i++)
        {
            _shellLifetime.Token.ThrowIfCancellationRequested();
            if (_shellWindows.Item(i) is not InternetExplorer window)
                continue;

            WindowInfo windowInfo;
            lock (_windowEntryDictLock)
            {
                if (_windowEntryDict.Keys.Contains(window))
                    continue;

                windowInfo = CreateWindowInfo(window);
                if (ExplorerWindowDiscovery.IsFileExplorerWindow(windowInfo.Identity.Handle) &&
                    !WinAPI.WinApi.IsWindowVisible(windowInfo.Identity.Handle))
                    continue;
                _windowEntryDict.Add(window, windowInfo);

            }

            if (_restoreTabs && !ExplorerWindowDiscovery.IsShownExplorerWindow(windowInfo.Identity.Handle))
                _preloadedSessionCandidates.TryAdd(windowInfo.Identity, (window, windowInfo));
            PreventWindowHiding(new IntPtr(window.HWND));

            if (MainWindowHandle == 0)
                MainWindowHandle = new IntPtr(window.HWND);

            registrations.Add(RegisterIndependentWindowAsync(window, windowInfo, windowInfo.Identity.Handle));
        }

        await Task.WhenAll(registrations);
        _shellLifetime.Token.ThrowIfCancellationRequested();
        _preExistingExplorerWindowsProtected = true;
        ScheduleShellWindowRegistration();
        ConcealPreloadedExplorerFrames();
        ObserveDesktopFolderOpen();
        if (_isForcingTabs)
            StartMergeSourceConcealPulse(500);
    }

    private string GetDefaultExplorerLocation()
    {
        var location = _getDefaultExplorerLaunchId() switch
        {
            2 => "shell:::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}",
            3 => "shell:::{088E3905-0323-4B02-9826-5D99428E115F}",
            4 => "shell:::{018D5C66-4533-4307-9B53-224DE2ED1FE6}",
            _ => "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
        };

        var pidl = _shellPathComparer.GetPidlFromPath(location);
        var path = ShellPathComparer.GetPathFromPidl(pidl);
        Marshal.FreeCoTaskMem(pidl);

        return Helper.NormalizeLocation(path ?? location);
    }

    private void DisposeShellObjects()
    {
        CompleteClosedSessions(shellEnded: true);
        _preExistingExplorerWindowsProtected = false;
        StopMergeSourceConcealPulse();
        ReleaseDesktopFolderOpenObserver();
        RecoverHiddenExplorerWindows("dispose-shell");
        var hookThread = Interlocked.Exchange(ref _winEventHookThread, null);
        hookThread?.Dispose();
        _eventObjectShowHookCallback = null;

        List<WindowEntry> windowEntries;
        lock (_windowEntryDictLock)
        {
            windowEntries = ((IEnumerable<WindowEntry>)_windowEntryDict).ToList();
            _windowEntryDict.Clear();
        }
        foreach (var (window, info) in windowEntries)
        {
            try
            {
                if (info.OnQuitHandler != null) window.OnQuit -= info.OnQuitHandler;
                if (info.OnNavigateHandler != null) window.NavigateComplete2 -= info.OnNavigateHandler;
                Marshal.ReleaseComObject(window);
                ExplorerWindowVisibility.ReleaseIdentityIfUntracked(info.Identity);
            }
            catch (Exception exception) when (exception is COMException or InvalidComObjectException)
            {
                ExplorerDebugLog.Write($"Shell window cleanup failed error={exception.GetType().Name}");
            }
        }

        if (_shellWindows != null)
        {
            try
            {
                if (_windowRegisteredHandler != null)
                    _shellWindows.WindowRegistered -= _windowRegisteredHandler;
                Marshal.ReleaseComObject(_shellWindows);
            }
            catch (Exception exception) when (exception is COMException or InvalidComObjectException)
            {
                ExplorerDebugLog.Write($"Shell connection cleanup failed error={exception.GetType().Name}");
            }
        }
        _windowRegisteredHandler = null;
        _shellWindows = null!;
        _shellPathComparer?.Dispose();
        _shellPathComparer = null!;
        lock (_closedWindowsLock)
            _closedWindows.Clear();
        ClearShellCaches();
        _processedHWnds.Clear();
        MainWindowHandle = 0;
    }

    private void ClearShellCaches()
    {
        _registrationRetryPending = false;
        _startupLocationCache.Clear();
        TabStrip.Clear();
        _hookedTopLevelUseCounts.Clear();
    }
}
