using System;
using System.Collections.Generic;
using System.Linq;
using WinTab.Helpers;

namespace WinTab.Hooks;

/// <param name="Tab">A still-current tab was moved to another window, not closed.</param>
/// <param name="Frame">The window the tab was closed in.</param>
internal readonly record struct ClosedTab(string Location, WindowIdentity Tab, WindowIdentity Frame,
    long Sequence = 0, long Generation = 0);

/// <summary>
/// Bounded, newest-first history of user-closed tabs. OnQuit and whole-window teardown can report the same
/// tab: bounded identity tombstones deduplicate both, even after the user has already restored that tab.
/// Different native tabs at the same location remain distinct records.
/// </summary>
internal sealed class ExplorerClosedTabHistory
{
    internal const int DefaultCapacity = 25;
    private readonly object _gate = new();
    private readonly int _capacity;
    private readonly LinkedList<ClosedTab> _tabs = new();
    private readonly HashSet<WindowIdentity> _seen = [];
    private readonly Queue<WindowIdentity> _seenOrder = new();
    private long _sequence;
    private long _generation;

    public ExplorerClosedTabHistory(int capacity = DefaultCapacity)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(capacity);
        _capacity = capacity;
    }

    public int Count
    {
        get { lock (_gate) return _tabs.Count; }
    }

    public void Push(ClosedTab tab)
    {
        if (string.IsNullOrWhiteSpace(tab.Location) || tab.Tab.Token == 0)
            return;
        lock (_gate)
        {
            if (!_seen.Add(tab.Tab))
                return;
            Remember(tab.Tab);
            _tabs.AddLast(tab with { Sequence = ++_sequence, Generation = _generation });
            Trim();
        }
    }

    /// <summary>Internal closes must also suppress a later whole-window fallback report.</summary>
    public void Ignore(WindowIdentity identity)
    {
        lock (_gate)
        {
            if (_seen.Add(identity))
                Remember(identity);
        }
    }

    private void Remember(WindowIdentity identity)
    {
        _seenOrder.Enqueue(identity);
        while (_seenOrder.Count > Math.Max(400, _capacity * 4))
            _seen.Remove(_seenOrder.Dequeue());
    }

    /// <summary>The remembered closes, newest first. A tab still alive elsewhere was moved, not closed, and is dropped.</summary>
    public ClosedTab[] Peek()
    {
        lock (_gate)
        {
            for (var node = _tabs.First; node != null;)
            {
                var next = node.Next;
                if (node.Value.Tab.IsCurrent)
                    _tabs.Remove(node);
                node = next;
            }
            return _tabs.Reverse().ToArray();
        }
    }

    /// <summary>Removes this entry of <see cref="Peek"/>; false when it was reopened, trimmed or forgotten meanwhile.</summary>
    public bool TryTake(ClosedTab tab)
    {
        lock (_gate)
        {
            if (tab.Generation != _generation)
                return false;
            for (var node = _tabs.Last; node != null; node = node.Previous)
            {
                if (node.Value.Sequence != tab.Sequence)
                    continue;
                _tabs.Remove(node);
                return true;
            }
            return false;
        }
    }

    /// <summary>A failed restore retains its original order, behind tabs closed while it was pending.</summary>
    public void Return(ClosedTab tab)
    {
        lock (_gate)
        {
            if (tab.Generation != _generation || tab.Sequence == 0)
                return; // Disabling the feature must also forget an in-flight restore's entry.
            var node = _tabs.Last;
            while (node != null && node.Value.Sequence > tab.Sequence)
                node = node.Previous;
            if (node?.Value.Sequence == tab.Sequence)
                return;
            if (node == null)
                _tabs.AddFirst(tab);
            else
                _tabs.AddAfter(node, tab);
            Trim();
        }
    }

    private void Trim()
    {
        while (_tabs.Count > _capacity)
            _tabs.RemoveFirst();
    }

    public void Clear()
    {
        lock (_gate)
        {
            _generation++;
            _tabs.Clear();
            _seen.Clear();
            _seenOrder.Clear();
        }
    }
}
