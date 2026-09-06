using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using WinTab.WinAPI;

namespace WinTab.Helpers;

public static class ExplorerWindowVisibility
{
    private sealed record VisibilityState(WindowIdentity Identity, bool WasLayered, uint ColorKey, byte Alpha, uint Flags)
    {
        public bool Recovering;
    }

    private static readonly ConcurrentDictionary<nint, VisibilityState> HiddenWindows = new();

    public static IEnumerable<nint> HiddenWindowHandles => HiddenWindows.Keys;

    public static bool Contains(nint hWnd) => HiddenWindows.TryGetValue(hWnd, out var state) && state.Identity.IsCurrent;

    public static bool Forget(nint hWnd) => HiddenWindows.TryRemove(hWnd, out _);

    internal static bool Forget(WindowIdentity identity) =>
        HiddenWindows.TryGetValue(identity.Handle, out var state) && state.Identity == identity &&
        HiddenWindows.TryRemove(new KeyValuePair<nint, VisibilityState>(identity.Handle, state));

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
            var wasLayered = (WinApi.GetWindowLong(hWnd, WinApi.GWL_EXSTYLE) & WinApi.WS_EX_LAYERED) != 0;
            uint colorKey = 0, flags = WinApi.LWA_ALPHA;
            byte alpha = 255;
            if (wasLayered && !WinApi.GetLayeredWindowAttributes(hWnd, out colorKey, out alpha, out flags))
                return;
            state = HiddenWindows.GetOrAdd(hWnd, new VisibilityState(identity, wasLayered, colorKey, alpha, flags));
        }
        lock (state)
        {
            if (state.Recovering || state.Identity != identity || !identity.IsCurrent)
                return;
            UpdateLayeredStyle(hWnd, remove: false);
            WinApi.SetLayeredWindowAttributes(hWnd, 0, 0, WinApi.LWA_ALPHA);
        }
    }

    public static int RestoreAll()
    {
        var restored = 0;
        foreach (var state in HiddenWindows.Values.ToArray())
        {
            if (Restore(state.Identity))
                restored++;
        }

        return restored;
    }

    public static bool Restore(nint hWnd, bool removeCache = true)
    {
        return HiddenWindows.TryGetValue(hWnd, out var state) && Restore(state.Identity, removeCache);
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
            if (state.WasLayered)
            {
                UpdateLayeredStyle(identity.Handle, remove: false);
                restored = WinApi.SetLayeredWindowAttributes(identity.Handle, state.ColorKey, state.Alpha, state.Flags);
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
