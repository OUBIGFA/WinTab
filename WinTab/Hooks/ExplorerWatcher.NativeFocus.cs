using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

public partial class ExplorerWatcher
{
    /// <summary>
    /// Open file location can focus an existing tab's file view without activating that tab or registering
    /// a new window. Follow the actual native keyboard focus, not changes to background file selections.
    /// </summary>
    private Task TryActivateNativeFocusedTabAsync(nint view)
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
        if (GetActiveTabHandle(parent) == tab || !HasNativeViewFocus(view, parent))
            return Task.CompletedTask;

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
            parentIdentity.IsCurrent && WinApi.GetParent(tab) == parent && HasNativeViewFocus(view, parent);

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
            using var operation = new MergeOperation(parentIdentity, generation, _shellLifetime.Token,
                IsOwnerCurrent, 1_000, _hookLifetime.Token);
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
                var selected = await SelectTabByHandle(parent, tab);
                ExplorerDebugLog.Write($"Native file-location focus selected={selected} hwnd={parent} tab={tab}");
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

    private static bool HasNativeViewFocus(nint view, nint parent)
    {
        var foreground = WinApi.GetForegroundWindow();
        if (foreground != parent && WinApi.GetAncestor(foreground, WinApi.GA_ROOT) != parent)
            return false;
        var thread = WinApi.GetWindowThreadProcessId(view, out _);
        var info = new GuiThreadInfo { Size = Marshal.SizeOf<GuiThreadInfo>() };
        return thread != 0 && GetGUIThreadInfo(thread, ref info) && info.Focus == view;
    }

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
