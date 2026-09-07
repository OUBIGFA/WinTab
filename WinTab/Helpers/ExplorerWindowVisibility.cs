using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using WinTab.WinAPI;

namespace WinTab.Helpers;

public static class ExplorerWindowVisibility
{
    private sealed record VisibilityState(WindowIdentity Identity, WindowVisibilitySnapshot Snapshot)
    {
        public bool Recovering;
    }

    private static readonly ConcurrentDictionary<nint, VisibilityState> HiddenWindows = new();

    public static IEnumerable<nint> HiddenWindowHandles => HiddenWindows.Keys;

    public static bool Contains(nint hWnd) => HiddenWindows.TryGetValue(hWnd, out var state) && state.Identity.IsCurrent;

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

    internal static void Hide(WindowIdentity identity)
    {
        if (!identity.IsCurrent)
            return;
        var hWnd = identity.Handle;
        if (HiddenWindows.TryGetValue(hWnd, out var stale) && !stale.Identity.IsCurrent)
            Forget(stale.Identity);
        if (!HiddenWindows.TryGetValue(hWnd, out var state))
        {
            var snapshot = WindowVisibilitySnapshot.Read(hWnd) ?? WindowVisibilitySnapshot.Capture(hWnd);
            if (snapshot == null)
                return;
            state = HiddenWindows.GetOrAdd(hWnd, new VisibilityState(identity, snapshot));
        }
        lock (state)
        {
            if (state.Recovering || state.Identity != identity || !identity.IsCurrent)
                return;
            if (!state.Snapshot.Save(hWnd))
            {
                Trace.TraceError($"Window recovery record could not be saved; the source was not hidden: {hWnd}");
                return;
            }
            UpdateLayeredStyle(hWnd, remove: false);
            WinApi.SetLayeredWindowAttributes(hWnd, 0, 0, WinApi.LWA_ALPHA);
        }
    }

    public static int RestoreAll()
    {
        var restored = 0;
        foreach (var handle in HiddenWindows.Keys.Concat(ExplorerWindowDiscovery.GetAllExplorerWindows()).Distinct().ToArray())
        {
            if (Restore(handle))
                restored++;
        }

        return restored;
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

    internal static bool Restore(WindowIdentity identity, bool removeCache = true)
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
            if (restored && removeCache)
                Forget(identity);
            return restored;
        }
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
