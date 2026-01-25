using System;
using System.Collections.Generic;
using System.Linq;
using WinTab.Helpers;
using WinTab.Models;
using WinTab.WinAPI;

namespace WinTab.Hooks;

internal sealed class ExplorerWindowHost : IWindowHost
{
    private readonly ExplorerWatcher _watcher;

    public ExplorerWindowHost(ExplorerWatcher watcher)
    {
        _watcher = watcher;
        _watcher.OnShellInitialized += HandleShellInitialized;
    }

    public WindowHostType HostType => WindowHostType.Explorer;
    public bool IsReady => _watcher.IsHookActive;

    public event Action<WindowEntry>? WindowCreated;
    public event Action<nint>? WindowDestroyed;
    public event Action<nint>? WindowActivated;

    public void Start() => _watcher.StartHook();
    public void Stop() => _watcher.StopHook();

    public bool TryActivate(nint hWnd)
    {
        if (!Helper.IsFileExplorerWindow(hWnd)) return false;
        WinApi.RestoreWindowToForeground(hWnd);
        return true;
    }

    public IReadOnlyCollection<WindowEntry> GetWindows()
    {
        var windows = _watcher.GetWindows();
        return windows.Select(w => new WindowEntry
        {
            Handle = w.Handle,
            Title = string.IsNullOrWhiteSpace(w.Name) ? w.DisplayLocation : w.Name,
            ProcessPath = string.Empty,
            HostType = WindowHostType.Explorer,
            State = WindowLifecycleState.Visible
        }).ToList();
    }

    private void HandleShellInitialized()
    {
        foreach (var window in GetWindows())
            WindowCreated?.Invoke(window);
    }

    public void Dispose()
    {
        _watcher.OnShellInitialized -= HandleShellInitialized;
    }
}


