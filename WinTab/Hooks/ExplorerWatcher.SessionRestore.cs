using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using SHDocVw;
using WinTab.Helpers;
using WinTab.Models;
using WinTab.WinAPI;

namespace WinTab.Hooks;

public partial class ExplorerWatcher
{
    /// <summary>
    /// The native half of restoration. A linked operation protects every command, including waits inside
    /// the existing tab-creation helpers. No foreground calls and no navigation of the initial tab occur.
    /// </summary>
    private sealed class NativeSessionRestore : IExplorerSessionRestoreEnvironment, IDisposable
    {
        private readonly ExplorerWatcher _owner;
        private readonly InternetExplorer _initialWindow;
        private readonly WindowInfo _initialInfo;
        private readonly WindowIdentity _frame;
        private readonly WindowIdentity _initial;
        private readonly string _initialLocation;
        private readonly int _generation;
        private readonly List<WindowIdentity> _expected;
        private readonly string _initialUiId;
        private Func<bool> _closeSelectedTab;
        private nint _expectedActive;
        private bool _creating;
        private bool _selecting;
        private bool _closingInitial;
        /// <summary>Closing the active placeholder hands activation to a restored tab Explorer chooses.</summary>
        private bool _awaitingSuccessor;

        public NativeSessionRestore(ExplorerWatcher owner, InternetExplorer window, WindowInfo info,
            string initialLocation, string initialUiId, int generation, CancellationToken cancellationToken)
        {
            _owner = owner;
            _initialWindow = window;
            _initialInfo = info;
            _frame = info.Identity;
            _initial = info.TabIdentity;
            _initialLocation = initialLocation;
            _generation = generation;
            _expected = [_initial];
            _expectedActive = _initial.Handle;
            _initialUiId = initialUiId;
            _closeSelectedTab = () => ExplorerTabAutomation.TryCloseSelectedTab(_frame.Handle, _initialUiId,
                () => _initial.IsCurrent && GetActiveTabHandle(_frame.Handle) == _initial.Handle && IsCurrent());
            Operation = new MergeOperation(_frame, generation, owner._shellLifetime.Token,
                IsCurrent, 60_000, cancellationToken);
        }

        public MergeOperation Operation { get; }
        public nint InitialTab => _initial.Handle;

        private bool IsCurrent()
        {
            var foreground = ExplorerNavigationAccess.ForegroundFrame();
            if (!_owner._restoreTabs || _owner._disposed || _generation != _owner._sessionGeneration || !_frame.IsCurrent ||
                !WinApi.IsWindowVisible(_frame.Handle) || foreground != _frame.Handle)
            {
                ExplorerDebugLog.Write($"Session native guard: frame or foreground changed hwnd={_frame.Handle} foreground={foreground} enabled={_owner._restoreTabs} generation={_generation}/{_owner._sessionGeneration}");
                return false;
            }
            var active = GetActiveTabHandle(_frame.Handle);
            if (!_selecting && active != _expectedActive &&
                !(_awaitingSuccessor && _expected.Any(tab => tab != _initial && tab.Handle == active)))
            {
                ExplorerDebugLog.Write($"Session native guard: active={active} expected={_expectedActive} selecting={_selecting}");
                return false;
            }
            if (!_closingInitial && (!_owner.IsCurrentWindow(_initialWindow, _initialInfo) ||
                !ExplorerSessionRestorePlan.SameLocation(_initialInfo.Location ?? string.Empty, _initialLocation)))
            {
                ExplorerDebugLog.Write("Session native guard: initial tab changed");
                return false;
            }
            if (_expected.Any(tab => !(_closingInitial && tab == _initial) &&
                (!tab.IsCurrent || WinApi.GetParent(tab.Handle) != _frame.Handle)))
            {
                ExplorerDebugLog.Write("Session native guard: expected tab moved or closed");
                return false;
            }

            var actual = ExplorerWindowDiscovery.GetStableExplorerTabs(_frame.Handle, ExplorerSession.MaxTabs + 2);
            if (actual == null)
                return false;
            var required = _expected.Where(tab => !(_closingInitial && tab == _initial)).Select(tab => tab.Handle).ToHashSet();
            if (!required.IsSubsetOf(actual))
                return false;
            var extra = actual.Count(tab => !required.Contains(tab) && !(_closingInitial && tab == _initial.Handle));
            return extra <= (_creating ? 1 : 0);
        }

        public void EnsureUnchanged()
        {
            Operation.ThrowIfInvalid();
            if (_owner.HasOtherShownSessionWindow(_frame.Handle) || Helper.IsCtrlShiftDown() ||
                !_closingInitial && (!IsSessionWindowIdle(_initialWindow) ||
                !ExplorerSessionRestorePlan.SameLocation(TryGetLocation(_initialWindow), _initialLocation)))
                throw new OperationCanceledException("The user changed the restoration window.");
        }

        /// <summary>
        /// Returns as soon as Explorer has created the tab. Its location is confirmed later together with the
        /// others, so the next tab is requested without waiting for this one's navigation.
        /// </summary>
        public async Task<nint> AppendTabAsync(string location)
        {
            EnsureUnchanged();
            _creating = true;
            try
            {
                var knownTabs = _expected.Select(tab => tab.Handle).ToArray();
                var handle = await _owner.CreateTabAtLocationAsync(_frame.Handle, _frame, knownTabs, location);
                if (handle == 0)
                    return 0;
                var identity = WindowIdentity.Capture(handle);
                _owner.EnsureWindowIdentity(identity);
                if (WinApi.GetParent(handle) != _frame.Handle)
                    throw new OperationCanceledException("The restored tab moved to another window.");
                _expected.Add(identity);
                return handle;
            }
            finally
            {
                _creating = false;
            }
        }

        public async Task<bool> ConfirmTabsAsync(IReadOnlyList<(nint Tab, string Location)> tabs)
        {
            EnsureUnchanged();
            ExplorerDebugLog.Write($"Session restore confirming {tabs.Count} tab locations hwnd={_frame.Handle}");
            var confirmations = await Task.WhenAll(tabs.Select(async tab =>
            {
                var window = await Helper.DoUntilNotDefaultAsync(
                    () => _owner.FindShellWindowByTabHandle(tab.Tab, _frame.Handle), NewTabWaitMs, 50, Operation.Token);
                var arrived = window != null && await _owner.WaitForNavigation(window, tab.Location, NavigationVerificationWaitMs);
                if (!arrived)
                    ExplorerDebugLog.Write($"Session restore tab unconfirmed tab={tab.Tab} shellWindow={window != null} target={tab.Location}");
                return arrived;
            }));
            EnsureUnchanged();
            var confirmed = confirmations.All(arrived => arrived);
            ExplorerDebugLog.Write($"Session restore locations confirmed={confirmed} hwnd={_frame.Handle}");
            return confirmed;
        }

        public async Task<bool> SelectTabAsync(nint tab)
        {
            EnsureUnchanged();
            var index = _expected.FindIndex(identity => identity.Handle == tab);
            if (index < 0)
                return false;
            _selecting = true;
            try
            {
                if (GetActiveTabHandle(_frame.Handle) != tab)
                {
                    _owner.SelectTabByIndex(_frame.Handle, index);
                    var selected = await Helper.DoUntilConditionAsync(() => GetActiveTabHandle(_frame.Handle),
                        active => active == tab, AppendedTabActivationWaitMs, 20, Operation.Token);
                    _owner.EnsureCurrentMerge();
                    if (selected != tab && !await _owner.SelectTabByHandle(_frame.Handle, tab, bringToFront: false))
                        return false;
                }
                _expectedActive = tab;
            }
            finally
            {
                _selecting = false;
            }
            // Observe native delayed focus restoration without calling SetForegroundWindow a second time.
            await Task.Delay(ForegroundSettleWaitMs, Operation.Token);
            EnsureUnchanged();
            return GetActiveTabHandle(_frame.Handle) == tab;
        }

        public async Task<bool> CloseInitialTabAsync()
        {
            EnsureUnchanged();
            // Explorer can apply a close sent to a background tab to its active tab, so the untouched
            // placeholder is closed only while it is itself the active tab (Explorer's own Ctrl+W command).
            if (_expected.Count < 2 || _expectedActive != InitialTab || GetActiveTabHandle(_frame.Handle) != InitialTab)
                return false;
            _closingInitial = true;
            _awaitingSuccessor = true;
            // UIA may wait for Explorer's OnQuit callback to return. Invoking it on the ShellWindows STA
            // would block that callback and can stall for UIA's minute-long timeout even after the tab closed.
            // The MTA worker still targets the exact initial tab's close button, never a queued Ctrl+W.
            ExplorerDebugLog.Write($"Session restore closing initial tab={InitialTab} hwnd={_frame.Handle}");
            if (!await Task.Run(_closeSelectedTab).ConfigureAwait(true))
                return false;
            ExplorerDebugLog.Write($"Session restore initial close invoked tab={InitialTab} hwnd={_frame.Handle}");
            var stillOpen = await Helper.DoUntilConditionAsync(() => _initial.IsCurrent,
                exists => !exists, FailedTabCleanupTimeoutMs, 20, Operation.Token);
            if (stillOpen)
                return false; // Never retry a posted close: a slow first request may still be in flight.
            _expected.Remove(_initial);
            var successor = await Helper.DoUntilConditionAsync(() => GetActiveTabHandle(_frame.Handle),
                active => _expected.Any(tab => tab.Handle == active), AppendedTabActivationWaitMs, 20, Operation.Token);
            if (!_expected.Any(tab => tab.Handle == successor))
                return false;
            _expectedActive = successor;
            _awaitingSuccessor = false;
            return true;
        }

        public void Dispose() => Operation.Dispose();
    }
}
