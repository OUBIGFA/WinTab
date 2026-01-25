using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using WinTab.Helpers;
using WinTab.Models;

namespace WinTab.Managers;

internal sealed class WindowGroupManager
{
    private readonly ConcurrentDictionary<Guid, WindowGroup> _groups = new();
    private readonly ConcurrentDictionary<nint, Guid> _windowToGroup = new();

    public WindowGroup CreateGroup(string name)
    {
        var group = new WindowGroup { Name = name };
        _groups[group.Id] = group;
        return group;
    }

    public bool TryAddToGroup(Guid groupId, WindowEntry entry)
    {
        if (!_groups.TryGetValue(groupId, out var group))
            return false;

        if (_windowToGroup.TryGetValue(entry.Handle, out var existing))
        {
            if (existing == groupId)
                return true;

            if (_groups.TryGetValue(existing, out var oldGroup))
                oldGroup.Members.RemoveAll(m => m.Handle == entry.Handle);
        }

        group.Members.RemoveAll(m => m.Handle == entry.Handle);
        group.Members.Add(entry);
        _windowToGroup[entry.Handle] = groupId;
        return true;
    }

    public bool RemoveFromGroup(nint hWnd)
    {
        if (!_windowToGroup.TryRemove(hWnd, out var groupId))
            return false;

        if (_groups.TryGetValue(groupId, out var group))
            group.Members.RemoveAll(m => m.Handle == hWnd);

        return true;
    }

    public IReadOnlyCollection<WindowGroup> GetGroups() => _groups.Values.ToList();

    public WindowGroup? GetGroup(Guid groupId) => _groups.TryGetValue(groupId, out var group) ? group : null;

    public Guid? GetGroupId(nint hWnd) => _windowToGroup.TryGetValue(hWnd, out var groupId) ? groupId : null;

    public void RemoveWindow(nint hWnd)
    {
        RemoveFromGroup(hWnd);
    }

    public void CleanupEmptyGroups()
    {
        foreach (var pair in _groups.ToArray())
        {
            if (pair.Value.Members.Count == 0)
                _groups.TryRemove(pair.Key, out _);
        }
    }
}


