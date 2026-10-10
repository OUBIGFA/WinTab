using System;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.WinAPI;

/// <summary>Stores opacity data on message-only handles, without changing any rendered surface.</summary>
internal sealed class TestWindowOpacity : IWindowOpacity
{
    internal static readonly TestWindowOpacity Instance = new();
    private const string State = "WinTab.Tests.Opacity";
    private const string Color = "WinTab.Tests.ColorKey";
    private const string Style = "WinTab.Tests.ExtendedStyle";
    private const string HasStyle = "WinTab.Tests.HasExtendedStyle";

    public int ReadStyle(nint handle) => WinApi.GetProp(handle, HasStyle) == 0
        ? WinApi.GetWindowLong(handle, WinApi.GWL_EXSTYLE) : unchecked((int)WinApi.GetProp(handle, Style));
    public void WriteStyle(nint handle, int style)
    {
        Check.That(!WinApi.IsWindowVisible(handle), "Background styles must never target a displayed window");
        Check.That(WinApi.SetProp(handle, Style, style) && WinApi.SetProp(handle, HasStyle, 1), "The message endpoint must retain its style data");
    }

    public bool TryRead(nint handle, out uint colorKey, out byte alpha, out uint flags)
    {
        var state = WinApi.GetProp(handle, State).ToInt64();
        colorKey = unchecked((uint)WinApi.GetProp(handle, Color).ToInt64());
        alpha = (byte)((state >> 2) & 255);
        flags = (uint)state & 3;
        return (ReadStyle(handle) & WinApi.WS_EX_LAYERED) != 0 && (state & 0x1000) != 0 && flags != 0;
    }

    public bool TryWrite(nint handle, uint colorKey, byte alpha, uint flags)
    {
        Check.That(!WinApi.IsWindowVisible(handle), "Background opacity tests must never operate on displayed windows");
        if (flags is 0 or > 3 || (ReadStyle(handle) & WinApi.WS_EX_LAYERED) == 0)
            return false;
        return WinApi.SetProp(handle, Color, unchecked((nint)(int)colorKey)) &&
            WinApi.SetProp(handle, State, (nint)(0x1000 | (alpha << 2) | (int)flags));
    }
}

/// <summary>Supplies only OS boundaries; ownership, persistence, retries and ordering remain production code.</summary>
internal static class BackgroundWindowVisibility
{
    public static void UpdateLayeredStyle(nint handle, bool remove) =>
        ExplorerWindowVisibility.UpdateLayeredStyle(handle, remove, TestWindowOpacity.Instance);
    public static void Hide(nint handle) => Hide(WindowIdentity.Capture(handle));
    public static void Hide(WindowIdentity identity) => _ = Hide(identity, _ => true);
    public static Task Hide(WindowIdentity identity, Func<nint, bool> remove) =>
        ExplorerWindowVisibility.Hide(identity, remove, TestWindowOpacity.Instance);
    public static bool Restore(nint handle, bool removeCache = true) =>
        ExplorerWindowVisibility.Restore(handle, removeCache, TestWindowOpacity.Instance, _ => true);
    public static bool Restore(WindowIdentity identity, bool removeCache = true) =>
        Restore(identity, removeCache, _ => true);
    public static bool Restore(WindowIdentity identity, bool removeCache, Func<nint, bool> add) =>
        ExplorerWindowVisibility.Restore(identity, removeCache, add);
}
