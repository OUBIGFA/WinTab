using System;
using System.Collections.Generic;
using System.Linq;
using WinTab.Helpers;
using WinTab.Models;

namespace WinTab.Hooks;

internal readonly record struct SessionTab(WindowIdentity Identity, string Location, string Title);
internal readonly record struct SessionVisualTab(string Id, string Title, bool Selected);

/// <summary>
/// Keeps complete snapshots independently of ShellWindows' per-tab OnQuit callbacks. UIA IDs are linked
/// to native tab identities through selection and unambiguous titles, never by activating tabs. If that
/// association is incomplete, every tab is retained and OrderVerified explicitly records the limitation.
/// </summary>
internal sealed class ExplorerSessionTracker
{
    internal const int CloseSettleMs = 750;
    private readonly object _gate = new();
    private readonly Dictionary<WindowIdentity, Frame> _frames = new();
    private long _observationOrder;

    /// <param name="inForeground">The window is in front now; the most recently used window is the live one.</param>
    public void Observe(WindowIdentity identity, SessionTab[] tabs, nint activeTab,
        SessionVisualTab[]? visualTabs, long now, bool inForeground = false)
    {
        if (tabs.Length > ExplorerSession.MaxTabs)
        {
            Forget(identity);
            return;
        }
        if (tabs.Length == 0 ||
            tabs.Any(tab => tab.Identity.Handle == 0 || string.IsNullOrWhiteSpace(tab.Location)) ||
            tabs.Select(tab => tab.Identity.Handle).Distinct().Count() != tabs.Length ||
            !tabs.Any(tab => tab.Identity.Handle == activeTab))
            return;
        if (visualTabs is { Length: > 0 } && visualTabs.Length != tabs.Length)
            return;
        // A cached strip may outlive a navigation. Linking stale titles to new native locations can
        // falsely verify a permutation of same-named folders, so keep the last order unverified.
        if (visualTabs != null && (!tabs.Select(tab => tab.Title).Order(StringComparer.OrdinalIgnoreCase)
                .SequenceEqual(visualTabs.Select(tab => tab.Title).Order(StringComparer.OrdinalIgnoreCase), StringComparer.OrdinalIgnoreCase) ||
            !tabs.Any(tab => tab.Identity.Handle == activeTab &&
                visualTabs.Any(item => item.Selected && StringComparer.OrdinalIgnoreCase.Equals(item.Title, tab.Title)))))
            visualTabs = null;

        lock (_gate)
        {
            if (!_frames.TryGetValue(identity, out var frame))
                _frames.Add(identity, frame = new Frame(++_observationOrder));
            if (inForeground)
                frame.LastUsed = now;
            frame.Observe(tabs, activeTab, visualTabs, now);
        }
    }

    /// <summary>The window came to the front; it may not have been captured yet.</summary>
    public void MarkUsed(WindowIdentity identity, long now)
    {
        lock (_gate)
            if (_frames.TryGetValue(identity, out var frame))
                frame.LastUsed = now;
    }

    /// <summary>
    /// The complete group of the most recently used window that <paramref name="qualifies"/>. A window never
    /// brought to the front ranks below every used one; among those the first tracked comes first.
    /// </summary>
    public (WindowIdentity Identity, ExplorerSession Session)? MostRecent(Func<ExplorerSession, bool> qualifies,
        Func<WindowIdentity, bool>? includeWindow = null)
    {
        lock (_gate)
        {
            foreach (var (identity, frame) in _frames.OrderByDescending(pair => pair.Value.LastUsed)
                .ThenBy(pair => pair.Value.Order))
            {
                if ((includeWindow == null || includeWindow(identity)) && frame.Snapshot(null) is { } session && qualifies(session))
                    return (identity, session);
            }
            return null;
        }
    }

    public void UpdateLocation(WindowIdentity identity, WindowIdentity tab, string location)
    {
        if (string.IsNullOrWhiteSpace(location)) return;
        lock (_gate)
        {
            if (!_frames.TryGetValue(identity, out var frame)) return;
            var index = Array.FindIndex(frame.Tabs, item => item.Identity == tab);
            if (index >= 0)
                frame.Tabs[index] = frame.Tabs[index] with { Location = location };
        }
    }

    public void TabClosing(WindowIdentity identity, WindowIdentity tab, long now)
    {
        lock (_gate)
        {
            if (_frames.TryGetValue(identity, out var frame) && frame.Tabs.Any(item => item.Identity == tab))
                frame.RemovalAt = now;
        }
    }

    /// <summary>A live tab moved to another window was not closed. Handle reuse alone does not identify a move.</summary>
    public ExplorerSession? WindowClosed(WindowIdentity identity, IReadOnlySet<WindowIdentity>? movedTabs = null)
    {
        lock (_gate)
            return _frames.Remove(identity, out var frame) ? frame.Snapshot(movedTabs) : null;
    }

    /// <summary>The window's settled tabs, with their identities and last locations.</summary>
    public SessionTab[] TabsOf(WindowIdentity identity)
    {
        lock (_gate)
            return _frames.TryGetValue(identity, out var frame) ? (SessionTab[])frame.Tabs.Clone() : [];
    }

    public void Forget(WindowIdentity identity)
    {
        lock (_gate) _frames.Remove(identity);
    }

    public void Clear()
    {
        lock (_gate) _frames.Clear();
    }

    private sealed class Frame(long order)
    {
        private readonly Dictionary<string, WindowIdentity> _visualOwners = new(StringComparer.Ordinal);
        public readonly long Order = order;
        public SessionTab[] Tabs = [];
        public long? RemovalAt;
        public long LastUsed;
        private WindowIdentity _active;
        private bool _orderVerified;
        private HashSet<WindowIdentity> _pendingSurvivors = [];
        private long _survivorsSince;

        public void Observe(SessionTab[] tabs, nint active, SessionVisualTab[]? visual, long now)
        {
            var identities = tabs.Select(tab => tab.Identity).ToHashSet();
            var lostTabs = Tabs.Any(tab => !identities.Contains(tab.Identity) &&
                !tabs.Any(current => current.Identity.Handle == tab.Identity.Handle));
            if (lostTabs)
            {
                RemovalAt ??= now;
                if (!_pendingSurvivors.SetEquals(identities))
                {
                    _pendingSurvivors = identities;
                    _survivorsSince = now;
                }
                // Window shutdown reports its tabs one by one. A surviving frame must settle before
                // its smaller group replaces the last complete snapshot; rapid shutdown keeps all tabs.
                if (now - RemovalAt.Value < CloseSettleMs || now - _survivorsSince < CloseSettleMs)
                    return;
            }
            RemovalAt = null;
            _pendingSurvivors = [];

            var previous = Tabs.Select(tab => tab.Identity).Where(identities.Contains).ToList();
            previous.AddRange(tabs.Select(tab => tab.Identity).Where(identity => !previous.Contains(identity)));
            var byIdentity = tabs.ToDictionary(tab => tab.Identity);
            var order = previous.ToArray();
            var activeIdentity = tabs.Single(tab => tab.Identity.Handle == active).Identity;
            var verified = tabs.Length == 1;
            if (visual is { Length: > 0 } && visual.All(tab => !string.IsNullOrEmpty(tab.Id)) &&
                visual.Select(tab => tab.Id).Distinct().Count() == visual.Length && visual.Count(tab => tab.Selected) == 1)
            {
                foreach (var key in _visualOwners.Keys.ToArray())
                {
                    var owner = _visualOwners[key];
                    var item = Array.Find(visual, item => item.Id == key);
                    if (!identities.Contains(owner) || item.Id == null ||
                        !StringComparer.OrdinalIgnoreCase.Equals(byIdentity[owner].Title, item.Title))
                        _visualOwners.Remove(key);
                }

                var selected = visual.Single(tab => tab.Selected);
                foreach (var key in _visualOwners.Where(pair => pair.Value == activeIdentity).Select(pair => pair.Key).ToArray())
                    _visualOwners.Remove(key);
                _visualOwners[selected.Id] = activeIdentity;

                var assigned = _visualOwners.Values.ToHashSet();
                foreach (var item in visual.Where(item => !_visualOwners.ContainsKey(item.Id)))
                {
                    if (string.IsNullOrEmpty(item.Title)) continue;
                    var matches = tabs.Where(tab => !assigned.Contains(tab.Identity) &&
                        StringComparer.OrdinalIgnoreCase.Equals(tab.Title, item.Title)).ToArray();
                    // Even duplicate locations can later navigate independently; do not permanently
                    // assign their UIA identities merely because their paths happen to match today.
                    if (matches.Length != 1) continue;
                    _visualOwners[item.Id] = matches[0].Identity;
                    assigned.Add(matches[0].Identity);
                }
                var unknownItems = visual.Where(item => !_visualOwners.ContainsKey(item.Id)).ToArray();
                var unknownTabs = previous.Where(identity => !assigned.Contains(identity)).ToArray();
                if (unknownItems.Length == 1 && unknownTabs.Length == 1)
                    _visualOwners[unknownItems[0].Id] = unknownTabs[0];

                verified = visual.All(item => _visualOwners.ContainsKey(item.Id)) &&
                    _visualOwners.Values.Distinct().Count() == tabs.Length;
                var remaining = new Queue<WindowIdentity>(previous.Where(identity => !_visualOwners.ContainsValue(identity)));
                order = visual.Select(item => _visualOwners.TryGetValue(item.Id, out var identity)
                    ? identity : remaining.Dequeue()).ToArray();
            }
            Tabs = order.Select(identity => byIdentity[identity]).ToArray();
            _active = activeIdentity;
            _orderVerified = verified;
        }

        public ExplorerSession? Snapshot(IReadOnlySet<WindowIdentity>? movedTabs)
        {
            var tabs = movedTabs == null ? Tabs : Tabs.Where(tab => !movedTabs.Contains(tab.Identity)).ToArray();
            if (tabs.Length == 0)
                return null;
            return new ExplorerSession
            {
                Locations = tabs.Select(tab => tab.Location).ToArray(),
                ActiveTabIndex = Math.Max(0, Array.FindIndex(tabs, tab => tab.Identity == _active)),
                OrderVerified = _orderVerified
            };
        }
    }
}
