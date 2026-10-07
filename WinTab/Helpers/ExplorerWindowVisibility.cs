using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using WinTab.WinAPI;

namespace WinTab.Helpers;

public static class ExplorerWindowVisibility
{
    private sealed record VisibilityState(WindowIdentity Identity, WindowVisibilitySnapshot Snapshot)
    {
        public bool Recovering;
        /// <summary>The taskbar button was removed for this concealment; it is not requested again on every re-hide.</summary>
        public bool ButtonRemoved;
        // Serializes only taskbar COM, never the immediate opacity change. Recovery marks Recovering
        // first, then waits here so a late DeleteTab cannot remove the restored window's button.
        public readonly object TaskbarGate = new();
        public bool ButtonRemovalPending;
        public Task ButtonRemoval = Task.CompletedTask;
        public Task<bool>? ButtonRestoration;
    }

    private static readonly ConcurrentDictionary<nint, VisibilityState> HiddenWindows = new();
    private static readonly ConcurrentDictionary<nint, Lazy<Task<bool>>> RecoveryWorkers = new();

    public static IEnumerable<nint> HiddenWindowHandles => HiddenWindows.Keys;

    public static bool Contains(nint hWnd) => HiddenWindows.TryGetValue(hWnd, out var state) && state.Identity.IsCurrent;

    // Retiring a Shell registration does not close its native window. A pending recovery still
    // needs the shared identity, including when restoring the taskbar button has failed.
    internal static void ReleaseIdentityIfUntracked(WindowIdentity identity)
    {
        if (!HiddenWindows.TryGetValue(identity.Handle, out var state) || state.Identity != identity)
            identity.Release();
    }

    public static bool Forget(nint hWnd) =>
        HiddenWindows.TryGetValue(hWnd, out var state) && Forget(state.Identity);

    internal static bool Forget(WindowIdentity identity)
    {
        if (!HiddenWindows.TryGetValue(identity.Handle, out var state) || state.Identity != identity)
            return false;
        lock (state)
        {
            state.Recovering = true;
            state.Snapshot.Remove(identity.Handle);
            return HiddenWindows.TryRemove(new KeyValuePair<nint, VisibilityState>(identity.Handle, state));
        }
    }

    public static void Hide(nint hWnd)
    {
        Hide(WindowIdentity.Capture(hWnd));
    }

    internal static void Hide(WindowIdentity identity) => _ = Hide(identity, TaskbarButton.Remove);

    internal static Task Hide(WindowIdentity identity, System.Func<nint, bool> removeButton)
    {
        if (!identity.IsCurrent)
            return Task.CompletedTask;
        var hWnd = identity.Handle;
        if (HiddenWindows.TryGetValue(hWnd, out var stale) && !stale.Identity.IsCurrent)
            Forget(stale.Identity);
        if (!HiddenWindows.TryGetValue(hWnd, out var state))
        {
            var snapshot = WindowVisibilitySnapshot.Read(hWnd) ?? WindowVisibilitySnapshot.Capture(hWnd);
            if (snapshot == null)
                return Task.CompletedTask;
            state = HiddenWindows.GetOrAdd(hWnd, new VisibilityState(identity, snapshot));
        }
        lock (state)
        {
            if (state.Recovering || state.Identity != identity || !identity.IsCurrent)
                return Task.CompletedTask;
            if (!state.Snapshot.Save(hWnd))
            {
                Trace.TraceError($"Window recovery record could not be saved; the source was not hidden: {hWnd}");
                return Task.CompletedTask;
            }
            UpdateLayeredStyle(hWnd, remove: false);
            WinApi.SetLayeredWindowAttributes(hWnd, 0, 0, WinApi.LWA_ALPHA);
            // Newer taskbars may acknowledge DeleteTab without removing the button. Exclude the
            // concealed frame natively as well; the persisted snapshot owns only these two style bits.
            SetTaskbarStyle(hWnd, 0x80); // WS_EX_TOOLWINDOW, without WS_EX_APPWINDOW
            // Never call Explorer's taskbar COM from the WinEvent thread, or hold the opacity lock
            // across it. A busy taskbar must not delay concealing this frame or the next one.
            if (!state.ButtonRemoved && !state.ButtonRemovalPending)
            {
                state.ButtonRemovalPending = true;
                state.ButtonRemoval = Task.Run(() => RemoveTaskbarButton(state, removeButton));
            }
            return state.ButtonRemoval;
        }
    }

    private static void RemoveTaskbarButton(VisibilityState state, System.Func<nint, bool> removeButton)
    {
        try
        {
            lock (state.TaskbarGate)
            {
                lock (state)
                    if (state.Recovering || !state.Identity.IsCurrent)
                        return;
                var removed = removeButton(state.Identity.Handle);
                lock (state)
                    state.ButtonRemoved = removed;
            }
        }
        catch (System.Exception exception)
        {
            Trace.TraceError($"Taskbar removal failed for {state.Identity.Handle}: {exception.GetType().Name}:{exception.Message}");
        }
        finally
        {
            lock (state)
                state.ButtonRemovalPending = false;
        }
    }

    public static int RestoreAll() => RestoreAll(HiddenWindows.Keys.Concat(ExplorerWindowDiscovery.GetAllExplorerWindows()));

    /// <summary>Restores the supplied native windows; callers can scope discovery to their own desktop.</summary>
    internal static int RestoreAll(IEnumerable<nint> handles)
    {
        var recoveries = handles.Distinct().Select(StartRecovery).ToArray();
        // SetWindowLong/ShowWindow can themselves wait for a foreign window thread. Start every
        // source independently before waiting, and keep unfinished work/records for a later retry.
        Task.WhenAll(recoveries).Wait(750);
        return recoveries.Count(task => task.IsCompletedSuccessfully && task.Result);
    }

    private static Task<bool> StartRecovery(nint handle)
    {
        var worker = RecoveryWorkers.GetOrAdd(handle, hWnd => new Lazy<Task<bool>>(() => Task.Run(() =>
        {
            try { return Restore(hWnd); }
            catch (Exception exception)
            {
                Trace.TraceError($"Window recovery failed for {hWnd}: {exception.GetType().Name}:{exception.Message}");
                return false;
            }
        })));
        var task = worker.Value;
        _ = task.ContinueWith(_ => RecoveryWorkers.TryRemove(new KeyValuePair<nint, Lazy<Task<bool>>>(handle, worker)),
            TaskScheduler.Default);
        return task;
    }

    public static bool Restore(nint hWnd, bool removeCache = true)
    {
        if (HiddenWindows.TryGetValue(hWnd, out var state))
            return Restore(state.Identity, removeCache);
        var snapshot = WindowVisibilitySnapshot.Read(hWnd);
        if (snapshot == null)
            return false;
        var identity = WindowIdentity.Capture(hWnd);
        if (!identity.IsCurrent)
            return false;
        HiddenWindows.TryAdd(hWnd, new VisibilityState(identity, snapshot));
        return Restore(identity, removeCache);
    }

    internal static bool Restore(WindowIdentity identity, bool removeCache = true) =>
        Restore(identity, removeCache, TaskbarButton.Add);

    internal static bool Restore(WindowIdentity identity, bool removeCache, System.Func<nint, bool> addButton)
    {
        if (!identity.IsCurrent)
        {
            Forget(identity);
            return false;
        }
        if (!HiddenWindows.TryGetValue(identity.Handle, out var state))
            return true;
        lock (state)
        {
            if (state.Identity != identity || !identity.IsCurrent)
                return false;
            state.Recovering = true;
            bool restored;
            var snapshot = state.Snapshot;
            var wasToolWindow = (WinApi.GetWindowLong(identity.Handle, WinApi.GWL_EXSTYLE) & 0x80) != 0;
            if (!SetTaskbarStyle(identity.Handle, snapshot.TaskbarStyle))
                return false;
            if (wasToolWindow && (snapshot.TaskbarStyle & 0x80) == 0 && WinApi.IsWindowVisible(identity.Handle))
            {
                // A frame first shown as a tool window may never have entered the taskbar's catalog.
                // Re-register it while still transparent, preserving its size/minimized state and never
                // activating it. Merely clearing WS_EX_TOOLWINDOW (or AddTab on newer shells) is not enough.
                WinApi.ShowWindow(identity.Handle, 0); // SW_HIDE
                WinApi.ShowWindow(identity.Handle, 8); // SW_SHOWNA, keep current placement without activation
                if (!WinApi.IsWindowVisible(identity.Handle)) return false;
            }
            if (snapshot.WasLayered)
            {
                UpdateLayeredStyle(identity.Handle, remove: false);
                restored = WinApi.SetLayeredWindowAttributes(identity.Handle, snapshot.ColorKey, snapshot.Alpha, snapshot.Flags);
            }
            else
            {
                WinApi.SetLayeredWindowAttributes(identity.Handle, 0, 255, WinApi.LWA_ALPHA);
                UpdateLayeredStyle(identity.Handle, remove: true);
                restored = (WinApi.GetWindowLong(identity.Handle, WinApi.GWL_EXSTYLE) & WinApi.WS_EX_LAYERED) == 0;
            }
            if (!restored)
                return false;
        }

        // Taskbar COM can stop answering. Keep one ordered worker per window and bound the caller's
        // wait, so a stalled DeleteTab/AddTab cannot strand every other transparent source at exit.
        Task<bool> recovery;
        lock (state)
        {
            if (state.ButtonRestoration == null || state.ButtonRestoration.IsCompleted)
                state.ButtonRestoration = Task.Run(() => RestoreTaskbarButton(state, removeCache, addButton));
            recovery = state.ButtonRestoration;
        }
        return recovery.Wait(100) && recovery.GetAwaiter().GetResult();
    }

    private static bool RestoreTaskbarButton(VisibilityState state, bool removeCache, System.Func<nint, bool> addButton)
    {
        var identity = state.Identity;
        try
        {
            return CompleteTaskbarRecovery(state, removeCache, addButton);
        }
        catch (System.Exception exception)
        {
            Trace.TraceError($"Taskbar recovery failed for {identity.Handle}: {exception.GetType().Name}:{exception.Message}");
            return false;
        }
    }

    private static bool CompleteTaskbarRecovery(VisibilityState state, bool removeCache, System.Func<nint, bool> addButton)
    {
        var identity = state.Identity;
        // The removal worker either sees Recovering and does nothing, or finishes DeleteTab first.
        // A delayed completion still checks the exact state and native identity before adding a button.
        lock (state.TaskbarGate)
        {
            // Another recovery may have completed while this one waited for COM. A newly concealed
            // lifetime of this same live HWND owns its own taskbar ordering; never issue a stale AddTab.
            if (!HiddenWindows.TryGetValue(identity.Handle, out var current))
                return identity.IsCurrent;
            if (!ReferenceEquals(current, state) || !identity.IsCurrent || !addButton(identity.Handle))
                return false;
            if (removeCache)
                Forget(identity);
            return true;
        }
    }

    private static bool SetTaskbarStyle(nint handle, int taskbarStyle)
    {
        var style = WinApi.GetWindowLong(handle, WinApi.GWL_EXSTYLE);
        var restored = (style & ~WindowVisibilitySnapshot.TaskbarStyleMask) | taskbarStyle;
        if (style != restored)
            WinApi.SetWindowLong(handle, WinApi.GWL_EXSTYLE, restored);
        return (WinApi.GetWindowLong(handle, WinApi.GWL_EXSTYLE) & WindowVisibilitySnapshot.TaskbarStyleMask) == taskbarStyle;
    }

    public static void UpdateLayeredStyle(nint hWnd, bool remove)
    {
        var exStyle = WinApi.GetWindowLong(hWnd, WinApi.GWL_EXSTYLE);
        var isLayered = (exStyle & WinApi.WS_EX_LAYERED) != 0;

        if (remove && isLayered)
            WinApi.SetWindowLong(hWnd, WinApi.GWL_EXSTYLE, exStyle & ~WinApi.WS_EX_LAYERED);

        if (!remove && !isLayered)
            WinApi.SetWindowLong(hWnd, WinApi.GWL_EXSTYLE, exStyle | WinApi.WS_EX_LAYERED);
    }

}
