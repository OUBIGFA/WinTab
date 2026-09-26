using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using SHDocVw;
using WinTab.Helpers;
using WinTab.Managers;
using WinTab.Models;
using WinTab.WinAPI;

namespace WinTab.Hooks;

using WindowEntry = DualKeyEntry<InternetExplorer, nint?, WindowInfo>;

public partial class ExplorerWatcher
{
    /// <summary>How old a foreground window's tab-strip snapshot may be when it is used, and when it is re-read.</summary>
    private const int SessionVisualFreshMs = 1_500;
    private const int SessionVisualRefreshMs = 1_000;
    // Explorer destroys HWNDs before its process exit is observable. Delay normal-close commits so a
    // crash's teardown callbacks cannot overwrite the already durable last-used window journal.
    internal const int SessionWindowCloseSettleMs = 1_000;
    private readonly ConcurrentDictionary<WindowIdentity, long> _sessionCloseObservedAt = new();
    private readonly ExplorerSessionTracker _sessionTracker = new();
    private readonly ExplorerSessionLocationPolicy _sessionLocationPolicy = new();
    private readonly ConcurrentDictionary<nint, WindowIdentity> _sessionWindows = new();
    private readonly ConcurrentDictionary<WindowIdentity, long> _sessionClosedAt = new();
    private readonly ConcurrentDictionary<nint, WindowIdentity> _sessionRestoreExclusions = new();
    private readonly ConcurrentDictionary<WindowIdentity, (InternetExplorer Window, WindowInfo Info)> _preloadedSessionCandidates = new();
    private readonly ConcurrentDictionary<WindowIdentity, SessionRestoreAttempt> _restoringSessionWindows = new();
    private readonly ExplorerClosedTabHistory _closedTabs = new();
    /// <summary>Tabs WinTab itself closes; their OnQuit is not a user closing a tab.</summary>
    private readonly ConcurrentDictionary<WindowIdentity, byte> _internalTabCloses = new();
    private readonly object _journalGate = new();
    private BackgroundRefreshCache<WindowIdentity, SessionVisualSnapshot>? _sessionVisuals;
    /// <summary>The last group: the last window closed, or the window in use when Explorer or Windows ended.</summary>
    private ExplorerSessionStore? _sessionStore;
    /// <summary>The window in use right now, journaled so that a shutdown or crash cannot lose it.</summary>
    private ExplorerSessionStore? _liveSessionStore;
    private ExplorerSession? _liveJournaled;
    private long _lastSaveStamp;
    private CancellationTokenSource _sessionLifetime = new();
    private CancellationToken _sessionCancellation;
    /// <summary>Capture runs whenever WinTab runs; it is what every kind of restore reads.</summary>
    private volatile bool _captureSessions;
    /// <summary>Automatic restoration into a newly opened sole window.</summary>
    private volatile bool _restoreTabs;
    private volatile bool _recordClosedTabs;
    private volatile bool _restoreOnAnyFolder;
    private volatile bool _restoreSingleTab;
    private int _sessionGeneration;
    private int _sessionRestoresInProgress;
    private long _sessionCloseSequence;

    private sealed record SessionVisualSnapshot(nint[] NativeTabs, nint ActiveTab, SessionVisualTab[] Tabs, long At);

    private sealed class SessionRestoreAttempt : IDisposable
    {
        private readonly CancellationTokenSource _cancellation;
        private readonly object _gate = new();
        private bool _sawForeground;
        private bool _watchNavigation;

        public SessionRestoreAttempt(CancellationToken lifetime, bool inForeground, WindowIdentity initialTab)
        {
            _cancellation = CancellationTokenSource.CreateLinkedTokenSource(lifetime);
            _sawForeground = inForeground;
            InitialTab = initialTab;
        }

        public WindowIdentity InitialTab { get; }
        public CancellationToken Token => _cancellation.Token;

        public void ObserveForeground(bool isTarget)
        {
            lock (_gate)
            {
                if (isTarget)
                    _sawForeground = true;
                else if (_sawForeground && !_cancellation.IsCancellationRequested)
                {
                    ExplorerDebugLog.Write("Session restore cancelled: foreground left the launch window");
                    _cancellation.Cancel();
                }
            }
        }

        public void ArmNavigation()
        {
            lock (_gate) _watchNavigation = true;
        }

        public void InitialTabNavigated()
        {
            lock (_gate)
                if (_watchNavigation && !_cancellation.IsCancellationRequested)
                {
                    ExplorerDebugLog.Write("Session restore cancelled: initial tab navigated");
                    _cancellation.Cancel();
                }
        }

        public void Dispose()
        {
            lock (_gate) _cancellation.Dispose();
        }
    }

    private void ObserveSessionForeground(nint foreground)
    {
        var frame = WinApi.GetAncestor(foreground, WinApi.GA_ROOT);
        foreach (var (identity, attempt) in _restoringSessionWindows)
            attempt.ObserveForeground(frame == identity.Handle);
        if (frame != 0 && _sessionWindows.TryGetValue(frame, out var used) && used.IsCurrent)
        {
            _sessionTracker.MarkUsed(used, Environment.TickCount64);
            _selectionWork.Request();
        }
    }

    /// <summary>
    /// Capture is the base of every restore and runs whatever the automatic-restore toggle says: the last
    /// closed window is saved when it closes, and the window in use is journaled as it changes, so that a
    /// shutdown, an Explorer crash or a power loss still leaves that window to restore. Capture and
    /// restoration are independent of merging.
    /// </summary>
    public void EnableSessionCapture()
    {
        if (_disposed || _captureSessions)
            return;
        _sessionStore ??= OpenSessionStore("session.json");
        _liveSessionStore ??= OpenSessionStore("session-live.json");
        _liveJournaled = _liveSessionStore.Snapshot;
        Interlocked.Exchange(ref _lastSaveStamp, Math.Max(_sessionStore.Snapshot?.SavedAt ?? 0, _liveJournaled?.SavedAt ?? 0));
        PromoteEndedLiveSession();
        _sessionVisuals ??= new BackgroundRefreshCache<WindowIdentity, SessionVisualSnapshot>(ReadSessionVisualSnapshot);
        _captureSessions = true;
        ExcludeExistingWindowsFromRestore();
        _selectionWork.Request();
    }

    /// <summary>
    /// The journaled window ended without closing when its Explorer process is gone: Windows shut down or
    /// restarted, Explorer crashed, or power was lost. It is then the last group, unless a window closed
    /// normally after it was journaled.
    /// </summary>
    private void PromoteEndedLiveSession()
    {
        var live = _liveSessionStore?.Snapshot;
        var saved = _sessionStore?.Snapshot;
        if (live?.Owner is not { } owner || live.SavedAt <= (saved?.SavedAt ?? 0) || owner.IsRunning())
            return;
        try
        {
            _ = _sessionStore!.SaveAsync(live with { Owner = null });
            ExplorerDebugLog.Write($"Session promoted from the journal of an ended Explorer tabs={live.Locations.Length}");
        }
        catch (Exception exception) when (exception is System.Text.Json.JsonException or ObjectDisposedException)
        {
            ReportStatus($"The window in use before Explorer ended could not be saved: {exception.Message}");
        }
    }

    /// <summary>Save times strictly increase, so the journal and the saved group can be ordered.</summary>
    private long NextSaveStamp()
    {
        while (true)
        {
            var last = Interlocked.Read(ref _lastSaveStamp);
            var next = Math.Max(DateTime.UtcNow.Ticks, last + 1);
            if (Interlocked.CompareExchange(ref _lastSaveStamp, next, last) == last)
                return next;
        }
    }

    private bool QualifiesAsGroup(ExplorerSession session) => _restoreSingleTab || session.Locations.Length >= 2;

    /// <summary>
    /// Journals the most recently used qualifying window when its group changed or another window took over.
    /// With no qualifying window the journal keeps its older entry, which the group saved later outdates.
    /// </summary>
    private void JournalLiveSession(bool force = false)
    {
        var store = _liveSessionStore;
        if (!_captureSessions || store == null)
            return;
        lock (_journalGate)
        {
            PromoteEndedLiveSession();
            // Keep the last intact journal while HWND destruction is being classified. In particular,
            // switching to a dying process's remaining window must not journal a shutdown intermediate.
            if (_sessionWindows.Values.Any(identity => !identity.IsCurrent))
                return;
            if (_sessionTracker.MostRecent(QualifiesAsGroup, identity => identity.IsCurrent) is not { } current)
                return;
            var (identity, session) = current;
            var processId = (int)identity.ProcessId;
            var journaled = _liveJournaled;
            if (!force && session.HasSameTabs(journaled) && journaled!.Owner?.ProcessId == processId)
                return;
            var owner = journaled?.Owner is { } known && known.ProcessId == processId && known.IsRunning()
                ? known : ExplorerSessionOwner.Of(processId);
            if (owner == null)
                return;
            var entry = session with { SavedAt = NextSaveStamp(), Owner = owner };
            try
            {
                _ = store.SaveAsync(entry);
                _liveJournaled = entry;
            }
            catch (Exception exception) when (exception is System.Text.Json.JsonException or ObjectDisposedException)
            {
                ExplorerDebugLog.Write($"The window in use could not be journaled: {exception.Message}");
            }
        }
    }

    /// <summary>WinTab or Windows is ending: journal the window in use and record nothing more.</summary>
    private void SuspendSessionCapture()
    {
        if (!_captureSessions)
            return;
        try { JournalLiveSession(); }
        catch (Exception exception) { ExplorerDebugLog.Write($"Final session journal failed: {exception.GetType().Name}"); }
        _captureSessions = false;
        // Do not put the only shutdown flush behind COM teardown, which may still be draining callbacks.
        var flushes = new[] { _sessionStore?.FlushAsync(), _liveSessionStore?.FlushAsync() }.OfType<Task<bool>>().ToArray();
        if (!Task.WhenAll(flushes).Wait(1_000))
            ExplorerDebugLog.Write("Final session flush is still pending; shutdown will not wait longer.");
    }

    /// <summary>
    /// Automatic restoration into a newly opened sole window. Restoration needs no drag observer: a window
    /// torn off another window never starts alone, and only a sole new window may restore.
    /// </summary>
    public void SetRestoreTabs(bool enabled)
    {
        if (_disposed || _restoreTabs == enabled)
            return;
        _restoreTabs = false;
        RenewSessionLifetime();
        _preloadedSessionCandidates.Clear();
        if (!enabled)
            return;
        EnableSessionCapture();
        ExcludeExistingWindowsFromRestore();
        // Explorer may already have registered a hidden preload before the toggle was enabled.
        // It becomes a new user window only when Explorer shows it later.
        lock (_windowEntryDictLock)
        {
            foreach (var entry in (IEnumerable<WindowEntry>)_windowEntryDict)
            {
                var info = entry.Value;
                if (IsCurrentWindow(entry.PrimaryKey, info) && info.TabIdentity.IsCurrent &&
                    !ExplorerWindowDiscovery.IsShownExplorerWindow(info.Identity.Handle))
                    _preloadedSessionCandidates.TryAdd(info.Identity, (entry.PrimaryKey, info));
            }
        }
        _restoreTabs = true;
        _selectionWork.Request();
    }

    public void SetRestoreOnAnyFolder(bool enabled)
    {
        if (_disposed || _restoreOnAnyFolder == enabled)
            return;
        _restoreOnAnyFolder = enabled;
        RenewSessionLifetime();
    }

    /// <summary>Off: a window with a single tab neither replaces the saved group nor is restored as one.</summary>
    public void SetRestoreSingleTab(bool enabled) => _restoreSingleTab = enabled;

    /// <summary>Off: closed tabs are not remembered and the remembered ones are forgotten.</summary>
    public void SetRecordClosedTabs(bool enabled)
    {
        _recordClosedTabs = enabled;
        if (!enabled)
            _closedTabs.Clear();
    }

    private void RenewSessionLifetime()
    {
        Interlocked.Increment(ref _sessionGeneration);
        _sessionLifetime.Cancel();
        _sessionLifetime.Dispose();
        _sessionLifetime = new CancellationTokenSource();
        _sessionCancellation = _sessionLifetime.Token;
    }

    private ExplorerSessionStore OpenSessionStore(string fileName)
    {
        var store = new ExplorerSessionStore(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WinTab", fileName));
        void ReportStorageError()
        {
            if (store.LastError is { } error)
                ReportStatus($"Explorer session storage is unavailable: {error.Message}");
        }
        store.ErrorChanged += ReportStorageError;
        ReportStorageError();
        return store;
    }

    private void ExcludeExistingWindowsFromRestore()
    {
        foreach (var handle in _getExplorerWindows().Where(ExplorerWindowDiscovery.IsShownExplorerWindow))
            ExcludeWindowFromSessionRestore(handle);
    }

    private void ExcludeWindowFromSessionRestore(nint handle)
    {
        if (!_captureSessions && !_restoreTabs) return;
        var identity = WindowIdentity.Capture(handle);
        if (identity.IsCurrent)
            _sessionRestoreExclusions[handle] = identity;
    }

    private async Task ObserveExplorerStateAsync()
    {
        await CacheActiveSelectionAsync();
        if (!_captureSessions)
            return;
        CaptureExplorerSessions();
        if (!_restoreTabs)
            return;
        foreach (var pair in _preloadedSessionCandidates.ToArray())
        {
            if (!pair.Key.IsCurrent)
            {
                _preloadedSessionCandidates.TryRemove(pair.Key, out _);
                continue;
            }
            if (!ExplorerWindowDiscovery.IsShownExplorerWindow(pair.Key.Handle))
                continue;
            _preloadedSessionCandidates.TryRemove(pair.Key, out _);
            await TryRestoreNewExplorerWindowAsync(pair.Value.Window, pair.Value.Info, wasPreloaded: true);
        }
    }

    private void CaptureExplorerSessions()
    {
        if (!_captureSessions || _disposed)
            return;
        var generation = _sessionGeneration;
        CompleteClosedSessions();
        WindowEntry[] entries;
        lock (_windowEntryDictLock)
            entries = ((IEnumerable<WindowEntry>)_windowEntryDict).ToArray();
        var foreground = ExplorerNavigationAccess.ForegroundFrame();
        foreach (var group in entries.Where(entry => !entry.Value.Closed).GroupBy(entry => entry.Value.Identity))
        {
            var identity = group.Key;
            if (!identity.IsCurrent || !ExplorerWindowDiscovery.IsShownExplorerWindow(identity.Handle) ||
                IsMergeSourceWindow(identity.Handle) || _restoringSessionWindows.ContainsKey(identity))
                continue;
            var native = ExplorerWindowDiscovery.GetStableExplorerTabs(identity.Handle, ExplorerSession.MaxTabs + 1);
            if (native is not { Length: > 0 })
                continue;
            if (native.Length > ExplorerSession.MaxTabs)
            {
                _sessionTracker.Forget(identity);
                _sessionWindows.TryRemove(new KeyValuePair<nint, WindowIdentity>(identity.Handle, identity));
                continue;
            }
            var owners = group.Where(entry => entry.OptionalKey is { } tab && IsCurrentTab(entry.Value, tab))
                .DistinctBy(entry => entry.OptionalKey).ToDictionary(entry => entry.OptionalKey!.Value);
            // Never save an incomplete ShellWindows catalog as if it were the whole window.
            if (native.Length != owners.Count || native.Any(tab => !owners.ContainsKey(tab)))
                continue;
            var tabs = new List<SessionTab>(native.Length);
            foreach (var handle in native)
            {
                var (window, info, _) = owners[handle];
                if (!IsCurrentWindow(window, info) || string.IsNullOrWhiteSpace(info.Location))
                    break;
                tabs.Add(new SessionTab(info.TabIdentity, info.Location, ReadSessionTitle(window, info)));
            }
            if (tabs.Count != native.Length || GetActiveTabHandle(identity.Handle) != native[0])
                continue;

            // Tabs are reordered only by the user in the foreground window. A background window's tab-strip
            // snapshot stays valid while its native tabs and selection are unchanged, so it is not re-read.
            var now = Environment.TickCount64;
            var inForeground = identity.Handle == foreground;
            var cached = _sessionVisuals!.TryGet(identity, out var snapshot) ? snapshot : null;
            var matches = cached != null && cached.ActiveTab == native[0] && cached.NativeTabs.ToHashSet().SetEquals(native);
            var visual = matches && (!inForeground || now - cached!.At <= SessionVisualFreshMs) ? cached!.Tabs : null;
            if (!matches || inForeground && now - cached!.At >= SessionVisualRefreshMs)
                _sessionVisuals.Request(identity);
            if (!_captureSessions || generation != _sessionGeneration)
                return;
            _sessionTracker.Observe(identity, tabs.ToArray(), native[0], visual, now, inForeground);
            _sessionWindows[identity.Handle] = identity;
        }
        foreach (var pair in _sessionRestoreExclusions)
            if (!pair.Value.IsCurrent)
                _sessionRestoreExclusions.TryRemove(pair);
        foreach (var pair in _internalTabCloses)
            if (!pair.Key.IsCurrent)
                _internalTabCloses.TryRemove(pair);
        JournalLiveSession();
    }

    /// <summary>A title only links a tab to the tab strip; it is read once per location, not on every pass.</summary>
    private static string ReadSessionTitle(InternetExplorer window, WindowInfo info)
    {
        var location = info.Location ?? string.Empty;
        if (info.SessionTitle is { } cached && cached.Location == location)
            return cached.Title;
        try
        {
            var title = window.LocationName ?? string.Empty;
            info.SessionTitle = (location, title);
            return title;
        }
        catch (COMException exception) when (!IsDisconnectedShell(exception))
        {
            // A title is optional. Losing it must not lose a real tab or block the other windows.
            return string.Empty;
        }
    }

    private static SessionVisualSnapshot? ReadSessionVisualSnapshot(WindowIdentity identity)
    {
        if (!identity.IsCurrent || !ExplorerWindowDiscovery.IsShownExplorerWindow(identity.Handle))
            return null;
        var before = ExplorerWindowDiscovery.GetStableExplorerTabs(identity.Handle, ExplorerSession.MaxTabs + 1);
        if (before is not { Length: > 0 } || before.Length > ExplorerSession.MaxTabs)
            return null;
        var tabs = ExplorerTabAutomation.ReadSessionTabs(identity.Handle);
        var after = ExplorerWindowDiscovery.GetStableExplorerTabs(identity.Handle, ExplorerSession.MaxTabs + 1);
        if (!identity.IsCurrent || after == null || !before.SequenceEqual(after) || tabs.Length != before.Length ||
            tabs.Any(tab => string.IsNullOrEmpty(tab.Id)))
            return null;
        return new SessionVisualSnapshot(before, before[0],
            tabs.Select(tab => new SessionVisualTab(tab.Id, tab.Title, tab.Selected)).ToArray(), Environment.TickCount64);
    }

    private void NotifySessionWindowDestroyed(nint handle)
    {
        if (_captureSessions && _sessionWindows.TryGetValue(handle, out var identity))
        {
            _sessionClosedAt.TryAdd(identity, Interlocked.Increment(ref _sessionCloseSequence));
            _sessionCloseObservedAt.TryAdd(identity, Environment.TickCount64);
            _selectionWork.Request();
        }
    }

    /// <summary>Fast close/reopen must wait for the closing window, not restore an older saved group.</summary>
    private async Task AwaitClosedSessionsAsync(CancellationToken token)
    {
        await Helper.DoUntilConditionAsync(() =>
        {
            CompleteClosedSessions();
            return _sessionWindows.Values.Any(identity => !identity.IsCurrent);
        }, pending => !pending, SessionWindowCloseSettleMs + 500, 25, token);
    }

    /// <param name="shellEnded">
    /// Diagnostic context for connection retirement. A COM disconnect alone is not proof of process exit;
    /// the persisted process incarnation determines whether the live journal takes precedence.
    /// </param>
    private void CompleteClosedSessions(bool shellEnded = false)
    {
        if (!_captureSessions || _sessionStore == null)
            return;
        var endedOwner = _liveJournaled?.Owner is { } owner && !owner.IsRunning() ? owner : null;
        PromoteEndedLiveSession();
        var now = Environment.TickCount64;
        var closed = _sessionWindows.Values.Where(identity => !identity.IsCurrent)
            .Select(identity => (Identity: identity, Order: _sessionClosedAt.GetOrAdd(identity,
                _ => Interlocked.Increment(ref _sessionCloseSequence))))
            .Where(item => endedOwner?.ProcessId == item.Identity.ProcessId ||
                now - _sessionCloseObservedAt.GetOrAdd(item.Identity, now) >= SessionWindowCloseSettleMs)
            .OrderBy(item => item.Order).ToArray();
        if (closed.Length == 0)
            return;
        var liveTabs = _getExplorerWindows().SelectMany(ExplorerWindowDiscovery.GetAllExplorerTabs)
            .Select(WindowIdentity.Read).Where(identity => identity.Token != 0).ToHashSet();
        var saved = false;
        foreach (var (identity, _) in closed)
        {
            if (endedOwner?.ProcessId != identity.ProcessId)
                RecordClosedWindowTabs(identity, liveTabs);
            var snapshot = _sessionTracker.WindowClosed(identity, liveTabs);
            _sessionWindows.TryRemove(new KeyValuePair<nint, WindowIdentity>(identity.Handle, identity));
            _sessionClosedAt.TryRemove(identity, out _);
            _sessionCloseObservedAt.TryRemove(identity, out _);
            _sessionVisuals?.Forget(identity);
            // The journal, not callback/dictionary order, owns recovery of an ended Explorer process.
            if (snapshot == null || endedOwner?.ProcessId == identity.ProcessId)
                continue;
            if (snapshot.Locations.Length < 2 && !_restoreSingleTab)
            {
                ExplorerDebugLog.Write("Session not saved: the closed window had a single tab; the saved group is kept");
                continue;
            }
            try
            {
                _ = _sessionStore.SaveAsync(snapshot with { SavedAt = NextSaveStamp() });
                saved = true;
            }
            catch (Exception exception) when (exception is System.Text.Json.JsonException or ObjectDisposedException)
            {
                ReportStatus($"The closed window's tabs could not be saved: {exception.Message}");
                continue;
            }
            ExplorerDebugLog.Write($"Session saved tabs={snapshot.Locations.Length} active={snapshot.ActiveTabIndex} orderVerified={snapshot.OrderVerified} shellEnded={shellEnded}");
        }
        // A window still in use must stay newer than the group just saved, or a later shutdown would lose it.
        if (saved)
            JournalLiveSession(force: true);
    }

    /// <summary>
    /// A user's tab is remembered as closed the moment Explorer reports it gone. Tabs of windows WinTab does
    /// not track as user windows (merge sources, windows being restored) and tabs WinTab closed are not.
    /// </summary>
    private void RecordClosedTab(WindowInfo info)
    {
        if (_internalTabCloses.TryRemove(info.TabIdentity, out _))
        {
            _closedTabs.Ignore(info.TabIdentity);
            return;
        }
        if (!_recordClosedTabs || !_captureSessions)
            return;
        var frame = info.Identity;
        if (!_sessionWindows.TryGetValue(frame.Handle, out var tracked) || tracked != frame ||
            _restoringSessionWindows.ContainsKey(frame) || string.IsNullOrWhiteSpace(info.Location))
            return;
        _closedTabs.Push(new ClosedTab(info.Location, info.TabIdentity, frame));
    }

    /// <summary>Explorer does not always report each tab of a closing window; the window's last group covers them.</summary>
    private void RecordClosedWindowTabs(WindowIdentity frame, IReadOnlySet<WindowIdentity> liveTabs)
    {
        if (!_recordClosedTabs)
            return;
        foreach (var tab in _sessionTracker.TabsOf(frame))
            if (!liveTabs.Contains(tab.Identity) && !_internalTabCloses.ContainsKey(tab.Identity))
                _closedTabs.Push(new ClosedTab(tab.Location, tab.Identity, frame));
    }

    private bool HasOtherShownSessionWindow(nint handle) =>
        _getExplorerWindows().Any(other => other != handle && ExplorerWindowDiscovery.IsShownExplorerWindow(other));

    /// <summary>
    /// Only the first registration/show of a genuinely new, sole window may restore. This is separate from
    /// the merge operation: disabling merging must not disable session capture or restoration. Until the
    /// attempt ends the window is not captured, so closing it early never replaces the saved group.
    /// </summary>
    private async Task<bool> TryRestoreNewExplorerWindowAsync(InternetExplorer window, WindowInfo info, bool wasPreloaded = false)
    {
        if (!_restoreTabs || _disposed || !IsCurrentWindow(window, info))
            return false;
        var identity = info.Identity;
        var handle = identity.Handle;
        if (_sessionRestoreExclusions.TryGetValue(handle, out var excluded) && excluded == identity)
            return false;
        if (!ExplorerWindowDiscovery.IsShownExplorerWindow(handle))
        {
            _preloadedSessionCandidates.TryAdd(identity, (window, info));
            return false;
        }
        // Consume this window's chance before awaiting anything; registration retries never restore twice.
        ExcludeWindowFromSessionRestore(handle);
        if ((!wasPreloaded && IsWindowProtected(handle)) || IsIndependentOpenRequested() || HasOtherShownSessionWindow(handle) ||
            HasOtherTrackedShellWindowForTopLevel(window, handle) || Volatile.Read(ref _sessionRestoresInProgress) != 0)
            return false;

        var generation = _sessionGeneration;
        var startedAt = System.Diagnostics.Stopwatch.GetTimestamp();
        string Elapsed() => $"{System.Diagnostics.Stopwatch.GetElapsedTime(startedAt).TotalMilliseconds:F0}ms";
        var attempt = new SessionRestoreAttempt(_sessionCancellation, ExplorerNavigationAccess.ForegroundFrame() == handle,
            info.TabIdentity);
        _restoringSessionWindows[identity] = attempt;
        _sessionTracker.Forget(identity);
        _sessionWindows.TryRemove(new KeyValuePair<nint, WindowIdentity>(handle, identity));
        try
        {
            var token = attempt.Token;
            await RestoreMergeSourceWindowAsync(handle);
            await RegisterIndependentWindowAsync(window, info, handle);
            if (!IsCurrentWindow(window, info) || !IsCurrentTab(info, info.TabIdentity.Handle))
                return true;
            await AwaitClosedSessionsAsync(token);
            var session = _sessionStore?.Snapshot;
            if (session == null)
                return true;
            if (session.Locations.Length < 2 && !_restoreSingleTab)
            {
                ExplorerDebugLog.Write($"Session not restored: the saved window has a single tab hwnd={handle}");
                return true;
            }
            ExplorerDebugLog.Write($"Session restore candidate hwnd={handle} saved={session.Locations.Length} registered at {Elapsed()}");
            var initialLocation = await _locationResolver.ResolveAsync(() => TryGetLocation(window),
                IsSessionStartLocation, token, isBusy: () => !IsSessionWindowIdle(window));
            var startLocation = !string.IsNullOrWhiteSpace(initialLocation) && IsSessionStartLocation(initialLocation);
            ExplorerDebugLog.Write($"Session restore launch location={initialLocation} start={startLocation} anyFolder={_restoreOnAnyFolder} resolved at {Elapsed()}");
            if (string.IsNullOrWhiteSpace(initialLocation) || !_restoreOnAnyFolder && !startLocation)
                return true;
            attempt.ArmNavigation();
            token.ThrowIfCancellationRequested();
            var available = await _sessionLocationPolicy.FindAvailableAsync(session, token);
            var plan = ExplorerSessionRestorePlan.Create(session, initialLocation, startLocation, _restoreOnAnyFolder,
                _restoreSingleTab, available);
            ExplorerDebugLog.Write($"Session restore plan tabs={plan?.Tabs.Length} closeInitial={plan?.CloseInitialTab} keepInitialActive={plan?.ActivateInitialTab} skipped={plan?.SkippedCount} at {Elapsed()}");
            if (plan == null)
                return true;
            if (plan.Tabs.Length == 0)
            {
                if (plan.SkippedCount > 0)
                    ReportStatus($"Session not restored; {plan.SkippedCount} unavailable or unsupported location(s) were skipped.");
                return true;
            }
            var foreground = await Helper.DoUntilConditionAsync(ExplorerNavigationAccess.ForegroundFrame,
                current => current == handle, 1_000, 20, token);
            if (!_restoreTabs || generation != _sessionGeneration || foreground != handle ||
                !IsCurrentWindow(window, info) || !IsSessionWindowIdle(window) || HasOtherShownSessionWindow(handle) ||
                !ExplorerSessionRestorePlan.SameLocation(initialLocation, TryGetLocation(window)) ||
                ExplorerWindowDiscovery.GetAllExplorerTabs(handle).Take(2).Count() != 1 ||
                GetActiveTabHandle(handle) != info.TabIdentity.Handle || IsIndependentOpenRequested())
            {
                ExplorerDebugLog.Write($"Session restore skipped: the window changed before restoring hwnd={handle} foreground={foreground} at {Elapsed()}");
                return true;
            }
            info.Location = initialLocation;
            await RestoreSessionInWindowAsync(window, info, initialLocation, plan, generation, token, requested: false);
            ExplorerDebugLog.Write($"Session restore finished hwnd={handle} at {Elapsed()}");
            return true;
        }
        catch (OperationCanceledException)
        {
            ExplorerDebugLog.Write($"Session restore cancelled before creation hwnd={handle}");
            return true;
        }
        catch (Exception exception) when (!IsDisconnectedShell(exception))
        {
            ReportStatus($"Session restoration stopped; the opened folder was retained ({exception.GetType().Name}: {exception.Message}).");
            return true;
        }
        finally
        {
            _restoringSessionWindows.TryRemove(identity, out _);
            attempt.Dispose();
            _selectionWork.Request();
        }
    }

    /// <summary>Explorer's start pages and the folder the user chose for a plain launch count as a normal launch.</summary>
    private bool IsSessionStartLocation(string location) =>
        ExplorerSessionLocationPolicy.IsStartPage(location) ||
        _defaultLocation is { Length: > 0 } defaultLocation &&
        StringComparer.OrdinalIgnoreCase.Equals(Helper.NormalizeLocation(location), defaultLocation);

    private static bool IsSessionWindowIdle(InternetExplorer window)
    {
        try { return !window.Busy; }
        catch (COMException exception) when (!IsDisconnectedShell(exception)) { return false; }
    }

    private async Task<string> ReadInitialSessionTabIdAsync(WindowInfo info, CancellationToken token)
    {
        var visuals = _sessionVisuals;
        if (visuals == null)
            return string.Empty;
        // UIA belongs on the existing bounded MTA workers, never the ShellWindows callback STA. Only a
        // strict-mode placeholder needs an identity for its close button; any-folder mode needs no read.
        visuals.Invalidate(info.Identity);
        var id = await Helper.DoUntilNotDefaultAsync(() =>
        {
            if (visuals.TryGet(info.Identity, out var snapshot) &&
                snapshot is { NativeTabs.Length: 1, Tabs: [{ Selected: true, Id.Length: > 0 } tab] } &&
                snapshot.NativeTabs[0] == info.TabIdentity.Handle && snapshot.ActiveTab == info.TabIdentity.Handle)
                return tab.Id;
            // A newly shown XAML tab strip may not have published its first item yet. The cache coalesces
            // retries while a read is running; a completed empty read must be requested again, not held forever.
            visuals.Request(info.Identity);
            return null;
        }, NewTabWaitMs, 20, token);
        return id ?? string.Empty;
    }

    /// <param name="requested">
    /// The user asked for this restore in a window WinTab opened for it; see <see cref="NativeSessionRestore"/>.
    /// </param>
    /// <returns>Whether every tab was restored and the saved active tab selected.</returns>
    private async Task<bool> RestoreSessionInWindowAsync(InternetExplorer window, WindowInfo info, string initialLocation,
        ExplorerSessionRestorePlan plan, int generation, CancellationToken token, bool requested)
    {
        await _toOpenWindowsLock.WaitAsync(token);
        var previous = _currentMerge.Value;
        NativeSessionRestore? environment = null;
        var guarded = false;
        try
        {
            var initialUiId = plan.CloseInitialTab ? await ReadInitialSessionTabIdAsync(info, token) : string.Empty;
            if (plan.CloseInitialTab && string.IsNullOrEmpty(initialUiId))
            {
                ReportStatus("Session restoration stopped: the initial tab's close button could not be identified safely.");
                return false;
            }
            environment = new NativeSessionRestore(this, window, info, initialLocation, initialUiId, generation, token, requested);
            Interlocked.Increment(ref _sessionRestoresInProgress);
            Interlocked.Increment(ref _tabSelectionsInProgress);
            guarded = true;
            _currentMerge.Value = environment.Operation;
            var result = await ExplorerSessionRestorer.RestoreAsync(plan, environment, environment.Operation.Token);
            ExplorerDebugLog.Write($"Session restore completed={result.Completed} restored={result.RestoredCount} skipped={plan.SkippedCount} closeInitial={plan.CloseInitialTab}");
            if (!result.Completed)
                ReportStatus($"Session restoration did not fully complete after {result.RestoredCount} tab(s); no further commands will be sent. Check the existing tabs.");
            else if (plan.SkippedCount > 0)
                ReportStatus($"Session restored; {plan.SkippedCount} unavailable or unsupported location(s) were skipped.");
            return result.Completed;
        }
        catch (OperationCanceledException)
        {
            ExplorerDebugLog.Write($"Session restore cancelled by navigation, focus, tabs or lifecycle hwnd={info.Identity.Handle} token={token.IsCancellationRequested} foreground={ExplorerNavigationAccess.ForegroundFrame()} active={GetActiveTabHandle(info.Identity.Handle)} initial={info.TabIdentity.Handle}");
            return false;
        }
        finally
        {
            _currentMerge.Value = previous;
            if (guarded)
            {
                Volatile.Write(ref _ignoreNativeFocusThrough, Environment.TickCount);
                Interlocked.Decrement(ref _tabSelectionsInProgress);
                Interlocked.Decrement(ref _sessionRestoresInProgress);
            }
            environment?.Dispose();
            _toOpenWindowsLock.Release();
        }
    }
}
