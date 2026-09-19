using System;
using System.Collections.Generic;
using System.Drawing;

namespace WinTab.Hooks;

/// <summary>
/// Recognises the user's own tab drags. When a tab is dragged off the tab row, Explorer itself opens a new
/// window for it; that window is what the user asked for and must not be merged back. Explorer keeps its
/// native drag untouched: this only watches where the left button went down and up and, once a window
/// appears shortly after such a drop, answers whether that window is the outcome of the drag.
/// </summary>
internal sealed class ExplorerTabTearOffTracker
{
    /// <summary>Explorer shows the window for a torn-off tab well within this time after the drop.</summary>
    internal const int WindowGraceMs = 3_000;

    private readonly IExplorerTabTearOffEnvironment _environment;
    private readonly object _gate = new();
    private DragCandidate? _candidate;
    private TearOff? _tearOff;

    public ExplorerTabTearOffTracker(IExplorerTabTearOffEnvironment environment)
    {
        _environment = environment;
    }

    /// <summary>A press on a window with several tabs may begin a tab drag; the decision waits for the drop.</summary>
    public void HandleLeftDown(Point point)
    {
        DragCandidate? candidate = null;
        if (_environment.IsEnabled)
        {
            var window = _environment.ResolveExplorerWindow(point);
            var tabCount = window == 0 ? 0 : _environment.CountTabs(window);
            if (tabCount >= 2)
            {
                var bounds = _environment.GetWindowBounds(window);
                if (!bounds.IsEmpty)
                {
                    candidate = new DragCandidate(window, tabCount, point, bounds, _environment.IsPointOnTab(point, window),
                        new HashSet<nint>(_environment.ListShownExplorerWindows()), _environment.DragWidth, _environment.DragHeight);
                }
            }
        }

        lock (_gate)
            _candidate = candidate;
    }

    /// <summary>
    /// The drop ends the drag. Explorer only tears the tab off when the press was on a tab title and the drop
    /// lands away from every tab row: a drop on the source row reorders the tabs or does nothing, a drop on
    /// another window's row moves the tab there, and a drag that moved the window was a title-bar drag.
    /// </summary>
    public void HandleLeftUp(Point point, long now)
    {
        DragCandidate? candidate;
        lock (_gate)
        {
            candidate = _candidate;
            _candidate = null;
        }

        if (candidate == null || !candidate.IsBeyondDragThreshold(point) || !_environment.IsEnabled)
            return;

        var source = candidate.Window;
        if (_environment.GetWindowBounds(source) != candidate.Bounds)
            return;

        // The title hit-test is filled in the background; a press it could not place yet is checked again now.
        if (!candidate.OnTab && !_environment.IsPointOnTab(candidate.Point, source))
            return;

        var dropWindow = _environment.ResolveExplorerWindow(point);
        if (dropWindow != 0 && IsOnTabRow(point, dropWindow, source, candidate.Bounds))
            return;

        lock (_gate)
            _tearOff = new TearOff(source, candidate.TabCount, _environment.TrackIdentity(source), candidate.ShownAtPress, now);
    }

    /// <summary>
    /// Explorer lays every window out alike, so a window whose own tab row is not known yet borrows the row
    /// offsets of the source window. When neither is known the drop counts as a row drop, which changes
    /// nothing: the window Explorer may open is merged as before.
    /// </summary>
    private bool IsOnTabRow(Point point, nint window, nint source, Rectangle sourceBounds)
    {
        var row = _environment.GetTabRow(window);
        if (row.IsEmpty)
        {
            var sourceRow = _environment.GetTabRow(source);
            var bounds = _environment.GetWindowBounds(window);
            if (sourceRow.IsEmpty || bounds.IsEmpty)
                return true;
            row = Rectangle.FromLTRB(bounds.Left, bounds.Top + (sourceRow.Top - sourceBounds.Top),
                bounds.Right, bounds.Top + (sourceRow.Bottom - sourceBounds.Top));
        }

        return row.Contains(point);
    }

    /// <summary>Whether a recent drop may still produce its window: nothing has been claimed and the grace period has not passed.</summary>
    public bool IsPending(long now)
    {
        lock (_gate)
            return _tearOff is { Claimed: 0 } tearOff && now - tearOff.At <= WindowGraceMs;
    }

    /// <summary>
    /// Whether <paramref name="window"/> is the window Explorer opened for the tab that was just dragged off.
    /// Explorer takes the tab out of the source window when it tears it off, so until the source has lost a
    /// tab no window can be the outcome of the drop. After that, the first single-tab window shown within
    /// the grace period that was not already on screen when the drag began is taken; the answer for that
    /// window then stays the same for as long as it exists, and every other window is refused.
    /// </summary>
    public bool TryClaim(nint window, long now)
    {
        if (window == 0 || !_environment.IsEnabled)
            return false;

        lock (_gate)
        {
            var tearOff = _tearOff;
            if (tearOff == null || window == tearOff.Source)
                return false;
            if (tearOff.Claimed == window)
                return tearOff.IsClaimedWindowAlive();
            if (tearOff.Claimed != 0 || now - tearOff.At > WindowGraceMs || tearOff.ShownAtPress.Contains(window))
                return false;

            var sourceTabCount = tearOff.IsSourceAlive() ? _environment.CountTabs(tearOff.Source) : 0;
            if (sourceTabCount >= tearOff.SourceTabCount)
                return false;
            if (!_environment.IsWindowShown(window) || _environment.CountTabs(window) != 1)
                return false;

            tearOff.Claim(window, _environment.TrackIdentity(window));
            return true;
        }
    }

    private sealed class DragCandidate(nint window, int tabCount, Point point, Rectangle bounds, bool onTab,
        HashSet<nint> shownAtPress, int dragWidth, int dragHeight)
    {
        public nint Window { get; } = window;
        public int TabCount { get; } = tabCount;
        public Point Point { get; } = point;
        public Rectangle Bounds { get; } = bounds;
        public bool OnTab { get; } = onTab;
        /// <summary>Windows already on screen when the drag began; none of them can be the window opened for it.</summary>
        public HashSet<nint> ShownAtPress { get; } = shownAtPress;

        public bool IsBeyondDragThreshold(Point current) =>
            Math.Abs(current.X - Point.X) > dragWidth || Math.Abs(current.Y - Point.Y) > dragHeight;
    }

    private sealed class TearOff(nint source, int sourceTabCount, Func<bool> isSourceAlive, HashSet<nint> shownAtPress, long at)
    {
        private Func<bool> _isClaimedWindowAlive = static () => false;

        public nint Source { get; } = source;
        public int SourceTabCount { get; } = sourceTabCount;
        public HashSet<nint> ShownAtPress { get; } = shownAtPress;
        public long At { get; } = at;
        public nint Claimed { get; private set; }

        /// <summary>A source that is gone has certainly lost its tab; a replacement window at its handle is not the source.</summary>
        public bool IsSourceAlive() => isSourceAlive();

        public void Claim(nint window, Func<bool> isAlive)
        {
            Claimed = window;
            _isClaimedWindowAlive = isAlive;
        }

        /// <summary>Explorer reuses window handles, so a later window at the same handle is not the claimed one.</summary>
        public bool IsClaimedWindowAlive() => _isClaimedWindowAlive();
    }
}

/// <summary>What the tracker needs to know about the screen; the mouse hook supplies the real answers.</summary>
internal interface IExplorerTabTearOffEnvironment
{
    bool IsEnabled { get; }
    int DragWidth { get; }
    int DragHeight { get; }

    /// <summary>The Explorer window under the point, or 0.</summary>
    nint ResolveExplorerWindow(Point point);

    /// <summary>Whether the point is on one of the window's tab titles.</summary>
    bool IsPointOnTab(Point point, nint explorerWindow);

    /// <summary>The band of the window that holds its tab titles, including the space beside them; empty while unknown.</summary>
    Rectangle GetTabRow(nint explorerWindow);

    /// <summary>The number of tabs in the window; 0 when the window is gone.</summary>
    int CountTabs(nint explorerWindow);

    /// <summary>Whether Explorer has shown the window; a frame it still keeps hidden cannot be the outcome of a drop.</summary>
    bool IsWindowShown(nint explorerWindow);

    /// <summary>Every Explorer window Explorer has shown, including those WinTab is concealing.</summary>
    IEnumerable<nint> ListShownExplorerWindows();

    /// <summary>The window's screen bounds, or empty when the window is gone.</summary>
    Rectangle GetWindowBounds(nint explorerWindow);

    /// <summary>Answers, later, whether the window at this handle is still the same window and not a replacement reusing the handle.</summary>
    Func<bool> TrackIdentity(nint explorerWindow);
}
