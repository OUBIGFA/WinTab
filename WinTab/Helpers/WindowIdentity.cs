using System;
using System.Threading;
using WinTab.WinAPI;

namespace WinTab.Helpers;

internal readonly record struct WindowIdentity(nint Handle, uint ProcessId, uint ThreadId, nint Token)
{
    private static readonly string PropertyName = "WinTab.WindowIdentity." + Guid.NewGuid().ToString("N");
    private static readonly object CaptureGate = new();
    private static int _nextToken;

    public bool IsCurrent => Handle != 0 && Token != 0 &&
        WinApi.GetProp(Handle, PropertyName) == Token &&
        WinApi.GetWindowThreadProcessId(Handle, out var processId) == ThreadId && processId == ProcessId;

    public static WindowIdentity Read(nint handle)
    {
        var threadId = WinApi.GetWindowThreadProcessId(handle, out var processId);
        return new WindowIdentity(handle, processId, threadId, WinApi.GetProp(handle, PropertyName));
    }

    public static WindowIdentity Capture(nint handle)
    {
        lock (CaptureGate)
        {
            var identity = Read(handle);
            if (identity.ThreadId == 0 || identity.IsCurrent)
                return identity;

            var token = (nint)Interlocked.Increment(ref _nextToken);
            return WinApi.SetProp(handle, PropertyName, token) ? identity with { Token = token } : default;
        }
    }

    public void Release()
    {
        lock (CaptureGate)
        {
            if (IsCurrent)
                WinApi.RemoveProp(Handle, PropertyName);
        }
    }
}
