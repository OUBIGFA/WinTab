using System;
using System.Collections.Generic;
using WinTab.Models;

namespace WinTab.Hooks;

internal interface IWindowHost : IDisposable
{
    WindowHostType HostType { get; }
    bool IsReady { get; }
    void Start();
    void Stop();
    IReadOnlyCollection<WindowEntry> GetWindows();
    bool TryActivate(nint hWnd);
    event Action<WindowEntry>? WindowCreated;
    event Action<nint>? WindowDestroyed;
    event Action<nint>? WindowActivated;
}


