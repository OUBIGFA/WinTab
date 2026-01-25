using System;
using System.Collections.Generic;
using WinTab.Hooks;
using WinTab.Models;

namespace WinTab.Managers;

internal sealed class WindowHostManager : IDisposable
{
    private readonly ExplorerWindowHost _explorerHost;
    private readonly GeneralWindowHost _generalHost;

    public event Action<WindowEntry>? WindowCreated;
    public event Action<nint>? WindowDestroyed;
    public event Action<nint>? WindowActivated;

    public WindowHostManager(ExplorerWatcher watcher)
    {
        _explorerHost = new ExplorerWindowHost(watcher);
        _generalHost = new GeneralWindowHost();

        _explorerHost.WindowCreated += entry => WindowCreated?.Invoke(entry);
        _generalHost.WindowCreated += entry => WindowCreated?.Invoke(entry);
        _generalHost.WindowDestroyed += hWnd => WindowDestroyed?.Invoke(hWnd);
        _generalHost.WindowActivated += hWnd => WindowActivated?.Invoke(hWnd);
    }

    public void StartGeneral() => _generalHost.Start();
    public void StopGeneral() => _generalHost.Stop();

    public IReadOnlyCollection<WindowEntry> GetExplorerWindows() => _explorerHost.GetWindows();
    public IReadOnlyCollection<WindowEntry> GetGeneralWindows() => _generalHost.GetWindows();

    public IReadOnlyCollection<WindowEntry> GetAllWindows()
    {
        var result = new List<WindowEntry>();
        result.AddRange(_explorerHost.GetWindows());
        result.AddRange(_generalHost.GetWindows());
        return result;
    }

    public bool TryActivateWindow(WindowEntry entry)
    {
        return entry.HostType switch
        {
            WindowHostType.Explorer => _explorerHost.TryActivate(entry.Handle),
            WindowHostType.General => _generalHost.TryActivate(entry.Handle),
            _ => false
        };
    }

    public void Dispose()
    {
        _generalHost.Dispose();
        _explorerHost.Dispose();
    }
}

