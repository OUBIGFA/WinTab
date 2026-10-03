using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

public partial class ExplorerWatcher
{
    private ExplorerNativeFocusActivation? _nativeForegroundActivation;

    private void ObserveNativeForegroundActivation(nint window, uint eventTime)
    {
        // The record is shared between the WinEvent and shell threads. Replacing the whole record also
        // invalidates queued requests when the user leaves this window or another activation occurs.
        var activation = CanReuseDesktopFolder && Volatile.Read(ref _tabSelectionsInProgress) == 0 &&
            ExplorerWindowDiscovery.IsFileExplorerWindow(window) && TryGetLastInputTime(out var inputTime)
            ? new ExplorerNativeFocusActivation(WindowIdentity.Capture(window), eventTime, inputTime)
            : null;
        Volatile.Write(ref _nativeForegroundActivation, activation);
    }

    /// <summary>
    /// Open file location can focus an existing tab's file view without activating that tab or registering
    /// a new window. Follow native focus notifications, including the brief toolbar handoff on activation,
    /// not changes to background file selections.
    /// </summary>
    private Task TryActivateNativeFocusedTabAsync(nint view, uint eventTime)
    {
        if (!CanReuseDesktopFolder || !WinApi.IsWindowHasClassName(view, "DirectUIHWND"))
            return Task.CompletedTask;
        var shellView = WinApi.GetParent(view);
        if (!WinApi.IsWindowHasClassName(shellView, "SHELLDLL_DefView"))
            return Task.CompletedTask;
        var tab = WinApi.GetParent(shellView);
        while (tab != 0 && !WinApi.IsWindowHasClassName(tab, "ShellTabWindowClass"))
            tab = WinApi.GetParent(tab);
        if (tab == 0)
            return Task.CompletedTask;
        var parent = WinApi.GetParent(tab);
        if (GetActiveTabHandle(parent) == tab)
            return Task.CompletedTask;

        var activation = Volatile.Read(ref _nativeForegroundActivation);
        var usedActivationHandoff = false;
        bool IsFocusCurrent()
        {
            if (HasNativeViewFocus(view, parent))
                return true;
            var followsActivation = activation != null && ReferenceEquals(activation, Volatile.Read(ref _nativeForegroundActivation)) &&
                TryGetNativeViewFocus(view, parent, out var focus) &&
                TryGetLastInputTime(out var inputTime) &&
                activation.CanFollowFocus(parent, focus, eventTime, unchecked((uint)Environment.TickCount), inputTime);
            usedActivationHandoff |= followsActivation;
            return followsActivation;
        }
        if (!IsFocusCurrent())
        {
            ExplorerDebugLog.Write($"Native file-location focus ignored hwnd={parent} tab={tab} reason=focus-moved");
            return Task.CompletedTask;
        }

        var viewIdentity = WindowIdentity.Capture(view);
        var tabIdentity = WindowIdentity.Capture(tab);
        var parentIdentity = WindowIdentity.Capture(parent);
        var requestedAt = Environment.TickCount64;
        var generation = _hookGeneration;
        var lastSelection = Volatile.Read(ref _ignoreNativeFocusThrough);
        bool IsCurrentRequest() => CanReuseDesktopFolder && generation == _hookGeneration &&
            Volatile.Read(ref _tabSelectionsInProgress) == 0 &&
            lastSelection == Volatile.Read(ref _ignoreNativeFocusThrough) &&
            Environment.TickCount64 - requestedAt < 1_000 && viewIdentity.IsCurrent && tabIdentity.IsCurrent &&
            parentIdentity.IsCurrent && WinApi.GetParent(tab) == parent && IsFocusCurrent();

        return RunShellWorkAsync(async () =>
        {
            if (!IsCurrentRequest()) return;
            var window = GetWindowByTabHandle(tab, parent);
            if (window == null || !TryGetTrackedEntry(window, out var entry) || !IsCurrentTab(entry.Value, tab))
                return;
            var info = entry.Value;
            bool IsOwnerCurrent() => CanReuseDesktopFolder && generation == _hookGeneration &&
                WinApi.GetAncestor(WinApi.GetForegroundWindow(), WinApi.GA_ROOT) == parent &&
                IsCurrentWindow(window, info) && IsCurrentTab(info, tab);
            // The request must still be fresh when it acquires the lock, but a valid request needs the
            // full 2.5-second selection budget. A one-second operation can stop on an intermediate tab.
            using var operation = new MergeOperation(parentIdentity, generation, _shellLifetime.Token,
                IsOwnerCurrent, 3_500, _hookLifetime.Token);
            var locked = false;
            var previousOperation = _currentMerge.Value;
            try
            {
                await _toOpenWindowsLock.WaitAsync(operation.Token);
                locked = true;
                // A focus event can be stale by the time the shell worker or another merge releases the lock.
                if (!IsCurrentRequest() || GetActiveTabHandle(parent) == tab) return;
                _currentMerge.Value = operation;
                operation.ThrowIfInvalid();
                var selected = await SelectTabByHandle(parent, tab, bringToFront: false);
                ExplorerDebugLog.Write($"Native file-location focus selected={selected} hwnd={parent} tab={tab} activationHandoff={usedActivationHandoff}");
            }
            catch (OperationCanceledException)
            {
                ExplorerDebugLog.Write($"Native file-location focus retired hwnd={parent} tab={tab}");
            }
            finally
            {
                _currentMerge.Value = previousOperation;
                if (locked) _toOpenWindowsLock.Release();
            }
        });
    }

    private static bool HasNativeViewFocus(nint view, nint parent) =>
        TryGetNativeViewFocus(view, parent, out var focus) && focus == view;

    private static bool TryGetNativeViewFocus(nint view, nint parent, out nint focus)
    {
        focus = 0;
        var foreground = WinApi.GetForegroundWindow();
        if (foreground != parent && WinApi.GetAncestor(foreground, WinApi.GA_ROOT) != parent)
            return false;
        var thread = WinApi.GetWindowThreadProcessId(view, out _);
        var info = new GuiThreadInfo { Size = Marshal.SizeOf<GuiThreadInfo>() };
        if (thread == 0 || !GetGUIThreadInfo(thread, ref info))
            return false;
        focus = info.Focus;
        return true;
    }

    private static bool TryGetLastInputTime(out uint time)
    {
        var info = new LastInputInfo { Size = (uint)Marshal.SizeOf<LastInputInfo>() };
        var read = GetLastInputInfo(ref info);
        time = info.Time;
        return read;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct LastInputInfo
    {
        public uint Size, Time;
    }

    [DllImport("user32.dll")]
    private static extern bool GetLastInputInfo(ref LastInputInfo info);

    [StructLayout(LayoutKind.Sequential)]
    private struct GuiThreadInfo
    {
        public int Size;
        public uint Flags;
        public nint Active, Focus, Capture, MenuOwner, MoveSize, Caret;
        public RECT CaretRect;
    }

    [DllImport("user32.dll")]
    private static extern bool GetGUIThreadInfo(uint thread, ref GuiThreadInfo info);
}
