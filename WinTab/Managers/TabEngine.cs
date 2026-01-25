using System;
using System.Collections.Generic;
using System.Linq;
using WinTab.Hooks;
using WinTab.Models;

namespace WinTab.Managers;

public sealed class TabEngine : IDisposable
{
    private readonly WindowHostManager _hostManager;
    private readonly WindowGroupManager _groupManager = new();
    private readonly List<WindowEntry> _cachedWindows = new();

    public event Action<WindowEntry>? WindowCreated;
    public event Action<nint>? WindowDestroyed;
    public event Action<nint>? WindowActivated;

    internal TabEngine(WindowHostManager hostManager)
    {
        _hostManager = hostManager;
        RefreshCache();
    }

    public IReadOnlyCollection<WindowEntry> GetAllWindows()
    {
        RefreshCache();
        return _cachedWindows.ToArray();
    }
    internal IReadOnlyCollection<WindowGroup> GetGroups() => _groupManager.GetGroups();

    internal WindowGroup CreateGroup(string name) => _groupManager.CreateGroup(name);

    internal bool AddToGroup(Guid groupId, WindowEntry entry) => _groupManager.TryAddToGroup(groupId, entry);

    public bool Activate(WindowEntry entry) => _hostManager.TryActivateWindow(entry);

    public void RemoveWindow(nint hWnd)
    {
        _groupManager.RemoveWindow(hWnd);
        _groupManager.CleanupEmptyGroups();
        _cachedWindows.RemoveAll(w => w.Handle == hWnd);
        WindowDestroyed?.Invoke(hWnd);
    }

    public void RefreshCache()
    {
        _cachedWindows.Clear();
        _cachedWindows.AddRange(_hostManager.GetAllWindows());
    }

    public void NotifyActivated(nint hWnd)
    {
        WindowActivated?.Invoke(hWnd);
    }

    public void NotifyCreated(WindowEntry entry)
    {
        _cachedWindows.RemoveAll(w => w.Handle == entry.Handle);
        _cachedWindows.Add(entry);
        WindowCreated?.Invoke(entry);
    }

    public void Dispose()
    {
        _hostManager.Dispose();
    }
}


