using SHDocVw;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Interop;
using WinTab.Models;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>
/// Keeps File Explorer in one window: new windows are merged into an existing one as tabs, an already open
/// folder is brought forward instead of opened twice, and tab groups are captured for restoration. This part
/// holds the shared state, the hook lifetime and the window event entry point; each concern lives in its own
/// partial file.
/// </summary>
[Fody.ConfigureAwait(true)]
public partial class ExplorerWatcher : IHook
{
    private static bool _instanceRunning;
    private static Guid _shellBrowserGuid = typeof(IShellBrowser).GUID;

    private ShellWindows _shellWindows = null!;
    private ShellPathComparer _shellPathComparer = null!;
    private readonly StaTaskScheduler _staTaskScheduler;
    private nint _mainWindowHandle;
    private readonly ConcurrentDictionary<nint, WindowIdentity> _processedHWnds = new();
    private readonly ConcurrentDictionary<nint, int> _hookedTopLevelUseCounts = new();
    private readonly ConcurrentDictionary<string, bool> _startupLocationCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly DualKeyDictionary<InternetExplorer, nint?, WindowInfo> _windowEntryDict = [];
    private readonly List<WindowRecord> _closedWindows = new();
    private readonly object _windowEntryDictLock = new(), _closedWindowsLock = new();
    private readonly SemaphoreSlim _toOpenWindowsLock = new(1);
    private readonly CoalescingAsyncWork _registrationWork;
    private readonly CoalescingAsyncWork _selectionWork;
    private readonly Func<int> _getDefaultExplorerLaunchId;
    private readonly Func<IEnumerable<nint>> _getExplorerWindows = ExplorerWindowDiscovery.GetAllExplorerWindows;
    private readonly ExplorerLaunchLocationResolver _locationResolver = new();
    private readonly ProcessWatcher _processWatcher;
    private int _mainExplorerProcessId;
    private Timer? _explorerCheckTimer;

    private WinEventHookThread? _winEventHookThread;
    private WinEventDelegate? _eventObjectShowHookCallback;
    private DShellWindowsEvents_WindowRegisteredEventHandler? _windowRegisteredHandler;

    private string _defaultLocation = null!;
    private volatile bool _reuseTabs = true;
    private volatile bool _isForcingTabs;
    private volatile bool _preExistingExplorerWindowsProtected;
    /// <summary>Explorer answered the direct tab request with a window or with nothing; the classic command is used instead.</summary>
    private volatile bool _directTabUnsupported;
    private readonly AsyncLocal<MergeOperation?> _currentMerge = new();
    private readonly MergeSourceConcealPulse _mergeSourceConcealPulse = new();
    private readonly ConcurrentDictionary<nint, ConcealedWindow> _mergeSourceHWnds = new();
    private readonly ConcurrentDictionary<nint, MergeOperation> _closingMergeSourceHWnds = new();
    private readonly ExplorerTabTearOffTracker _tabTearOff;
    private readonly ExplorerTabTearOffHook _tabTearOffHook;
    private readonly ConcurrentDictionary<nint, bool> _tornOffTabReleaseWatches = new();

    internal TabStripHitTester TabStrip { get; } = new();
    public bool IsHookActive => _isForcingTabs;
    public bool IsShellReady => _mainExplorerProcessId != 0 && _shellWindows != null;
    public event Action? OnShellInitialized;

    public ExplorerWatcher(Func<int>? getDefaultExplorerLaunchId = null)
    {
        if (_instanceRunning)
            throw new InvalidOperationException("Only one instance of ExplorerWatcher is allowed at a time.");
        _instanceRunning = true;

        _staTaskScheduler = new StaTaskScheduler();
        _registrationWork = new CoalescingAsyncWork(() => RunShellWorkAsync(ProcessRegisteredShellWindowsAsync));
        _selectionWork = new CoalescingAsyncWork(() => RunShellWorkAsync(ObserveExplorerStateAsync));
        _mergeSafetyTimer = new Timer(RecoverExpiredMergeSources, null, Timeout.Infinite, Timeout.Infinite);
        _selectionTimer = new Timer(state => _selectionWork.Request(), null, Timeout.Infinite, Timeout.Infinite);
        _getDefaultExplorerLaunchId = getDefaultExplorerLaunchId ?? (static () => 1);
        _tabTearOff = new ExplorerTabTearOffTracker(new ExplorerTabTearOffEnvironment(TabStrip, () => _isForcingTabs));
        _tabTearOffHook = new ExplorerTabTearOffHook(_tabTearOff);
        _processWatcher = new ProcessWatcher("explorer");
        _processWatcher.ProcessTerminated += OnExplorerProcessTerminated;
        StartExplorerProcessCheck();
    }

    public void StartHook()
    {
        if (_isForcingTabs || _disposed) return;
        RecoverHiddenExplorerWindows("start-hook");
        Interlocked.Increment(ref _hookGeneration);
        _hookLifetime.Dispose();
        _hookLifetime = new CancellationTokenSource();
        _isForcingTabs = true;
        StartTabTearOffHook();
        ScheduleShellWindowRegistration();
        ConcealPreloadedExplorerFrames();
        ExplorerDebugLog.Write("StartHook");
    }

    public void StopHook()
    {
        _isForcingTabs = false;
        Interlocked.Increment(ref _hookGeneration);
        _hookLifetime.Cancel();
        _tabTearOffHook.StopHook();
        StopMergeSourceConcealPulse();
        RecoverHiddenExplorerWindows("stop-hook");
    }

    /// <summary>Merging keeps working without the drag observer; only dragged-off tabs would then be merged back.</summary>
    private void StartTabTearOffHook()
    {
        try
        {
            _tabTearOffHook.StartHook();
        }
        catch (Exception exception)
        {
            ReportStatus($"Tab drag detection is unavailable ({exception.GetType().Name}); tabs dragged out of a window may be merged back.");
        }
    }
    public void SetReuseTabs(bool reuseTabs) => _reuseTabs = reuseTabs;

    private void OnWindowShown(nint hWinEventHook, uint eventType, nint hWnd, int idObject, int idChild, uint dwEventThread, uint dWmsEventTime)
    {
        if (_captureSessions && !_disposed && eventType == WinApi.EVENT_SYSTEM_FOREGROUND)
            ObserveSessionForeground(hWnd);
        if (_captureSessions && !_disposed && eventType == WinApi.EVENT_OBJECT_DESTROY && idObject == 0 && idChild == 0)
        {
            NotifySessionWindowDestroyed(hWnd);
            return;
        }
        if ((!_isForcingTabs && !_captureSessions) || !_preExistingExplorerWindowsProtected || _disposed || hWnd == 0) return;

        if (eventType == WinApi.EVENT_OBJECT_FOCUS)
        {
            // Explorer may report only an accessible item's focus when the hidden file view already
            // owns keyboard focus (not another OBJID_CLIENT/CHILDID_SELF event). The handler validates
            // the native file-view ancestry and live keyboard focus, so item notifications are safe too.
            if (_isForcingTabs && Volatile.Read(ref _tabSelectionsInProgress) == 0 &&
                unchecked((int)dWmsEventTime - Volatile.Read(ref _ignoreNativeFocusThrough)) > 0)
            {
                try { _ = TryActivateNativeFocusedTabAsync(hWnd); }
                catch (Exception exception) when (exception is ObjectDisposedException or InvalidOperationException or TaskSchedulerException)
                {
                    ExplorerDebugLog.Write($"Native file-location focus could not be scheduled: {exception.Message}");
                }
            }
            return;
        }

        // OBJID_WINDOW = 0 and CHILDID_SELF = 0. The system-wide WinEvent hook range fires for every
        // accessibility sub-element on the desktop (caret, focus, menu items, scrollbars, list items,
        // alerts, etc.). Without this filter every caret blink in every program would push us into the
        // expensive Explorer-top-level lookup + ShellWindows registration scheduler + 25ms-pulse worker.
        if (idObject != 0 || idChild != 0) return;

        // Skip non-Explorer events entirely so the conceal pulse and registration scheduler only fire
        // when something actually touched a CabinetWClass window.
        var explorerTopLevel = GetExplorerTopLevelWindow(hWnd);
        if (explorerTopLevel == 0) return;
        if (_captureSessions)
        {
            if (IsIndependentOpenRequested())
                ExcludeWindowFromSessionRestore(explorerTopLevel);
            // Explorer showed a (possibly preloaded) frame or added a tab: a launch to check or a group to capture.
            _selectionWork.Request();
        }

        TryHideIncomingExplorerWindow(explorerTopLevel);
        StartMergeSourceConcealPulse();
        ScheduleShellWindowRegistration();
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        var startedAt = Environment.TickCount64;
        // Windows may be ending the session with Explorer windows still open; they must survive in the journal.
        SuspendSessionCapture();
        _disposed = true;
        _isForcingTabs = false;
        _preExistingExplorerWindowsProtected = false;
        Interlocked.Increment(ref _hookGeneration);
        Interlocked.Increment(ref _shellGeneration);
        _shellLifetime.Cancel();
        _explorerCheckTimer?.Dispose();
        _hookLifetime.Cancel();
        _mergeSafetyTimer.Dispose();
        _selectionTimer.Dispose();
        _sessionLifetime.Cancel();
        _sessionVisuals?.Dispose();
        StopMergeSourceConcealPulse();
        _tabTearOffHook.Dispose();
        Interlocked.Exchange(ref _winEventHookThread, null)?.Dispose();
        RecoverHiddenExplorerWindows("dispose");
        TabStrip.Dispose();
        _processWatcher.ProcessTerminated -= OnExplorerProcessTerminated;
        _processWatcher.Dispose();

        var cleanup = Task.Run(FinishDisposalAsync);
        var remaining = Math.Max(0, 2_000 - (int)(Environment.TickCount64 - startedAt));
        if (!cleanup.Wait(remaining))
            ExplorerDebugLog.Write("Shell cleanup is still pending; shutdown will not wait longer.");
        GC.SuppressFinalize(this);
    }
}
