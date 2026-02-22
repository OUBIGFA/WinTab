using WinTab.Core.Interfaces;
using WinTab.Core.Models;

namespace WinTab.TabHost;

/// <summary>
/// Manages all active tab groups. Implements <see cref="IGroupManager"/> to provide
/// a central registry of groups and window-to-group mappings, and reacts to
/// window lifecycle events (e.g., destruction) to keep state consistent.
/// </summary>
public sealed class TabGroupManager : IGroupManager, IDisposable
{
    private readonly IWindowManager _windowManager;
    private readonly IWindowEventSource _windowEventSource;
    private readonly int _tabBarHeight;

    /// <summary>All active tab group controllers keyed by group ID.</summary>
    private readonly Dictionary<Guid, TabGroupController> _controllers = [];

    /// <summary>Lookup from window handle to the group ID it belongs to.</summary>
    private readonly Dictionary<IntPtr, Guid> _windowToGroup = [];

    private readonly object _lock = new();
    private bool _disposed;

    // ─── Events (IGroupManager) ─────────────────────────────────────────

    public event EventHandler<TabGroup>? GroupCreated;
    public event EventHandler<TabGroup>? GroupDisbanded;
    public event EventHandler<TabGroup>? TabSwitched;
    public event EventHandler<TabGroup>? TabAdded;
    public event EventHandler<TabGroup>? TabRemoved;

    // ─── Constructor ────────────────────────────────────────────────────

    /// <summary>
    /// Creates a new <see cref="TabGroupManager"/>.
    /// </summary>
    /// <param name="windowManager">Window manipulation service.</param>
    /// <param name="windowEventSource">System-wide window event source.</param>
    /// <param name="tabBarHeight">Height of overlay tab bars in pixels.</param>
    public TabGroupManager(IWindowManager windowManager, IWindowEventSource windowEventSource, int tabBarHeight = 32)
    {
        _windowManager = windowManager ?? throw new ArgumentNullException(nameof(windowManager));
        _windowEventSource = windowEventSource ?? throw new ArgumentNullException(nameof(windowEventSource));
        _tabBarHeight = tabBarHeight;

        // Auto-remove windows that are destroyed.
        _windowEventSource.WindowDestroyed += OnWindowDestroyed;
    }

    // ─── IGroupManager Implementation ───────────────────────────────────

    /// <inheritdoc />
    public TabGroup CreateGroup(IntPtr window1, IntPtr window2)
    {
        if (window1 == IntPtr.Zero) throw new ArgumentException("Invalid window handle.", nameof(window1));
        if (window2 == IntPtr.Zero) throw new ArgumentException("Invalid window handle.", nameof(window2));

        lock (_lock)
        {
            // Ensure neither window is already in a group.
            if (_windowToGroup.ContainsKey(window1))
                throw new InvalidOperationException($"Window {window1} is already in a tab group.");
            if (_windowToGroup.ContainsKey(window2))
                throw new InvalidOperationException($"Window {window2} is already in a tab group.");

            var controller = new TabGroupController(_windowManager, window1, window2, _tabBarHeight);

            // Wire up controller events.
            controller.TabAdded += OnControllerTabAdded;
            controller.TabRemoved += OnControllerTabRemoved;
            controller.ActiveTabChanged += OnControllerActiveTabChanged;
            controller.Disbanded += OnControllerDisbanded;

            _controllers[controller.GroupId] = controller;
            _windowToGroup[window1] = controller.GroupId;
            _windowToGroup[window2] = controller.GroupId;

            var snapshot = controller.ToSnapshot();
            GroupCreated?.Invoke(this, snapshot);
            return snapshot;
        }
    }

    /// <inheritdoc />
    public TabGroup? AddToGroup(Guid groupId, IntPtr window)
    {
        if (window == IntPtr.Zero) return null;

        lock (_lock)
        {
            if (!_controllers.TryGetValue(groupId, out var controller))
                return null;

            // Don't add if already in any group.
            if (_windowToGroup.ContainsKey(window))
                return null;

            var vm = controller.AddTab(window);
            if (vm is null) return null;

            _windowToGroup[window] = groupId;

            var snapshot = controller.ToSnapshot();
            TabAdded?.Invoke(this, snapshot);
            return snapshot;
        }
    }

    /// <inheritdoc />
    public bool RemoveFromGroup(IntPtr window)
    {
        if (window == IntPtr.Zero) return false;

        lock (_lock)
        {
            if (!_windowToGroup.TryGetValue(window, out var groupId))
                return false;

            if (!_controllers.TryGetValue(groupId, out var controller))
            {
                _windowToGroup.Remove(window);
                return false;
            }

            _windowToGroup.Remove(window);
            bool removed = controller.RemoveTab(window);

            if (removed)
            {
                var snapshot = controller.ToSnapshot();
                TabRemoved?.Invoke(this, snapshot);
            }

            return removed;
        }
    }

    /// <inheritdoc />
    public TabGroup? GetGroupForWindow(IntPtr window)
    {
        lock (_lock)
        {
            if (!_windowToGroup.TryGetValue(window, out var groupId))
                return null;

            if (!_controllers.TryGetValue(groupId, out var controller))
                return null;

            return controller.ToSnapshot();
        }
    }

    /// <inheritdoc />
    public IReadOnlyList<TabGroup> GetAllGroups()
    {
        lock (_lock)
        {
            return _controllers.Values
                .Select(c => c.ToSnapshot())
                .ToList()
                .AsReadOnly();
        }
    }

    /// <inheritdoc />
    public bool SwitchTab(Guid groupId, int tabIndex)
    {
        lock (_lock)
        {
            if (!_controllers.TryGetValue(groupId, out var controller))
                return false;

            bool switched = controller.SwitchToTab(tabIndex);

            if (switched)
            {
                var snapshot = controller.ToSnapshot();
                TabSwitched?.Invoke(this, snapshot);
            }

            return switched;
        }
    }

    /// <inheritdoc />
    public bool CloseTab(Guid groupId, IntPtr window)
    {
        lock (_lock)
        {
            if (!_controllers.TryGetValue(groupId, out var controller))
                return false;

            _windowToGroup.Remove(window);
            return controller.CloseTab(window);
        }
    }

    /// <inheritdoc />
    public bool MoveTab(Guid groupId, IntPtr window, int offset)
    {
        lock (_lock)
        {
            if (!_controllers.TryGetValue(groupId, out var controller))
                return false;

            return controller.MoveTab(window, offset);
        }
    }

    /// <inheritdoc />
    public bool RenameGroup(Guid groupId, string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return false;

        lock (_lock)
        {
            if (!_controllers.TryGetValue(groupId, out var controller))
                return false;

            controller.Name = name;
            return true;
        }
    }

    /// <inheritdoc />
    public bool DisbandGroup(Guid groupId)
    {
        lock (_lock)
        {
            if (!_controllers.TryGetValue(groupId, out var controller))
                return false;

            var snapshot = controller.ToSnapshot();

            // Remove all window mappings for this group.
            foreach (var tab in controller.Tabs)
                _windowToGroup.Remove(tab.Handle);

            // Unhook events and disband.
            UnwireController(controller);
            controller.DisbandGroup();
            _controllers.Remove(groupId);

            GroupDisbanded?.Invoke(this, snapshot);
            return true;
        }
    }

    // ─── Public Helpers ─────────────────────────────────────────────────

    /// <summary>
    /// Gets the <see cref="TabGroupController"/> for a given group ID.
    /// Used by advanced consumers that need direct controller access.
    /// </summary>
    public TabGroupController? GetController(Guid groupId)
    {
        lock (_lock)
        {
            return _controllers.GetValueOrDefault(groupId);
        }
    }

    /// <summary>
    /// Checks whether a window is currently in any tab group.
    /// </summary>
    public bool IsWindowGrouped(IntPtr window)
    {
        lock (_lock)
        {
            return _windowToGroup.ContainsKey(window);
        }
    }

    // ─── Controller Event Handlers ──────────────────────────────────────

    private void OnControllerTabAdded(object? sender, TabItemViewModel vm)
    {
        // The window-to-group mapping is already updated in AddToGroup.
    }

    private void OnControllerTabRemoved(object? sender, TabItemViewModel vm)
    {
        lock (_lock)
        {
            _windowToGroup.Remove(vm.Handle);
        }
    }

    private void OnControllerActiveTabChanged(object? sender, TabItemViewModel vm)
    {
        if (sender is TabGroupController controller)
        {
            var snapshot = controller.ToSnapshot();
            TabSwitched?.Invoke(this, snapshot);
        }
    }

    private void OnControllerDisbanded(object? sender, EventArgs e)
    {
        if (sender is TabGroupController controller)
        {
            lock (_lock)
            {
                // Clean up all mappings for this group.
                var keysToRemove = _windowToGroup
                    .Where(kvp => kvp.Value == controller.GroupId)
                    .Select(kvp => kvp.Key)
                    .ToList();

                foreach (var key in keysToRemove)
                    _windowToGroup.Remove(key);

                UnwireController(controller);
                _controllers.Remove(controller.GroupId);
            }

            var snapshot = controller.ToSnapshot();
            GroupDisbanded?.Invoke(this, snapshot);
        }
    }

    // ─── Window Event Handlers ──────────────────────────────────────────

    private void OnWindowDestroyed(object? sender, IntPtr hwnd)
    {
        // When a window is destroyed, remove it from its group.
        lock (_lock)
        {
            if (!_windowToGroup.TryGetValue(hwnd, out var groupId))
                return;

            if (!_controllers.TryGetValue(groupId, out var controller))
            {
                _windowToGroup.Remove(hwnd);
                return;
            }

            _windowToGroup.Remove(hwnd);

            // Use Dispatcher to ensure we're on the UI thread for WPF operations.
            if (controller.Overlay is not null)
            {
                controller.Overlay.Dispatcher.BeginInvoke(() =>
                {
                    controller.RemoveTab(hwnd);
                });
            }
            else
            {
                controller.RemoveTab(hwnd);
            }
        }
    }

    // ─── Private Helpers ────────────────────────────────────────────────

    private void UnwireController(TabGroupController controller)
    {
        controller.TabAdded -= OnControllerTabAdded;
        controller.TabRemoved -= OnControllerTabRemoved;
        controller.ActiveTabChanged -= OnControllerActiveTabChanged;
        controller.Disbanded -= OnControllerDisbanded;
    }

    // ─── IDisposable ────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _windowEventSource.WindowDestroyed -= OnWindowDestroyed;

        lock (_lock)
        {
            foreach (var controller in _controllers.Values)
            {
                UnwireController(controller);

                // Restore all grouped windows before disposing internals.
                controller.DisbandGroup();
                controller.Dispose();
            }

            _controllers.Clear();
            _windowToGroup.Clear();
        }
    }
}
