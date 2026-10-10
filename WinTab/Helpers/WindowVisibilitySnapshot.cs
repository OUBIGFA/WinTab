using System;
using System.Diagnostics;
using WinTab.WinAPI;

namespace WinTab.Helpers;

internal sealed record WindowVisibilitySnapshot(bool WasLayered, uint ColorKey, byte Alpha, uint Flags, nint Token, int TaskbarStyle)
{
    internal const int TaskbarStyleMask = 0x00040080; // WS_EX_APPWINDOW | WS_EX_TOOLWINDOW
    private const string TaskbarStyleProperty = "WinTab.HiddenWindow.TaskbarStyle.v1";
    private const string StateProperty = "WinTab.HiddenWindow.State.v1";
    private const string ColorKeyProperty = "WinTab.HiddenWindow.ColorKey.v1";
    private const string TokenProperty = "WinTab.HiddenWindow.Token.v1";
    private const int StateMarker = 0x1000;

    public static WindowVisibilitySnapshot? Capture(nint handle, IWindowOpacity? opacity = null)
    {
        opacity ??= NativeWindowOpacity.Instance;
        var style = opacity.ReadStyle(handle);
        var wasLayered = (style & WinApi.WS_EX_LAYERED) != 0;
        uint colorKey = 0, flags = WinApi.LWA_ALPHA;
        byte alpha = 255;
        if (wasLayered && (!opacity.TryRead(handle, out colorKey, out alpha, out flags) ||
            flags == 0 || (flags & ~3u) != 0))
        {
            Trace.TraceError($"Could not capture window opacity: {handle}");
            return null;
        }
        return new WindowVisibilitySnapshot(wasLayered, colorKey, alpha, flags, Random.Shared.Next(1, int.MaxValue),
            style & TaskbarStyleMask);
    }

    public static WindowVisibilitySnapshot? Read(nint handle, IWindowOpacity? opacity = null)
    {
        var token = WinApi.GetProp(handle, TokenProperty);
        var state = WinApi.GetProp(handle, StateProperty).ToInt64();
        if (token == 0 || (state & ~0x7ffL) != StateMarker)
            return null;
        var colorKey = unchecked((uint)WinApi.GetProp(handle, ColorKeyProperty).ToInt64());
        var flags = (uint)(state >> 9) & 3;
        if (flags == 0)
            return null;
        var savedStyle = WinApi.GetProp(handle, TaskbarStyleProperty).ToInt64();
        // Older concealments did not alter these style bits, so their current bits are the original ones.
        var taskbarStyle = (savedStyle & ~TaskbarStyleMask) == 1
            ? (int)savedStyle & TaskbarStyleMask : (opacity ?? NativeWindowOpacity.Instance).ReadStyle(handle) & TaskbarStyleMask;
        if (WinApi.GetProp(handle, TokenProperty) != token)
            return null;
        return new WindowVisibilitySnapshot((state & 1) != 0, colorKey, (byte)((state >> 1) & 255), flags, token, taskbarStyle);
    }

    public bool Save(nint handle)
    {
        var existingToken = WinApi.GetProp(handle, TokenProperty);
        if (existingToken != 0)
            return existingToken == Token && WinApi.SetProp(handle, TaskbarStyleProperty, TaskbarStyle | 1);
        var state = StateMarker | (WasLayered ? 1 : 0) | (Alpha << 1) | ((int)Flags << 9);
        if (WinApi.SetProp(handle, TaskbarStyleProperty, TaskbarStyle | 1) &&
            WinApi.SetProp(handle, ColorKeyProperty, unchecked((nint)(int)ColorKey)) &&
            WinApi.SetProp(handle, StateProperty, state) && WinApi.SetProp(handle, TokenProperty, Token))
            return true;
        WinApi.RemoveProp(handle, StateProperty);
        WinApi.RemoveProp(handle, ColorKeyProperty);
        WinApi.RemoveProp(handle, TaskbarStyleProperty);
        return false;
    }

    public void Remove(nint handle)
    {
        if (WinApi.GetProp(handle, TokenProperty) != Token)
            return;
        WinApi.RemoveProp(handle, TokenProperty);
        WinApi.RemoveProp(handle, StateProperty);
        WinApi.RemoveProp(handle, ColorKeyProperty);
        WinApi.RemoveProp(handle, TaskbarStyleProperty);
    }
}
