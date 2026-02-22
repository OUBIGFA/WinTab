using System.Collections.ObjectModel;
using WinTab.Core.Interfaces;
using WinTab.Core.Models;
using WinTab.Platform.Win32;

namespace WinTab.TabHost;

/// <summary>
/// Manages a single tab group: the overlay tab bar, tab switching, window
/// show/hide, and position tracking for a set of grouped windows.
/// </summary>
public sealed class TabGroupController : IDisposable
{
    private readonly IWindowManager _windowManager;
    private readonly OverlayTracker _overlayTracker;
    private bool _disposed;

    // ─── Properties ─────────────────────────────────────────────────────

    /// <summary>Unique identifier for this tab group.</summary>
    public Guid GroupId { get; }

    /// <summary>Human-readable group name.</summary>
    public string Name { get; set; } = "Default";

    /// <summary>Observable collection of tabs in this group.</summary>
    public ObservableCollection<TabItemViewModel> Tabs { get; } = [];

    /// <summary>The currently active (visible, foreground) tab.</summary>
    public TabItemViewModel? ActiveTab { get; private set; }

    /// <summary>The overlay window displaying the tab bar.</summary>
    public TabOverlayWindow? Overlay { get; private set; }

    /// <summary>The timestamp when this group was created.</summary>
    public DateTime CreatedAt { get; }

    // ─── Events ─────────────────────────────────────────────────────────

    /// <summary>Raised after the active tab changes.</summary>
    public event EventHandler<TabItemViewModel>? ActiveTabChanged;

    /// <summary>Raised when a tab is added.</summary>
    public event EventHandler<TabItemViewModel>? TabAdded;

    /// <summary>Raised when a tab is removed.</summary>
    public event EventHandler<TabItemViewModel>? TabRemoved;

    /// <summary>Raised when the group is disbanded (all tabs restored).</summary>
    public event EventHandler? Disbanded;

    // ─── Constructor ────────────────────────────────────────────────────

    /// <summary>
    /// Creates a new tab group from two initial windows.
    /// </summary>
    /// <param name="windowManager">Window manipulation service.</param>
    /// <param name="window1">First window to include (becomes the active tab).</param>
    /// <param name="window2">Second window to include (hidden initially).</param>
    /// <param name="tabBarHeight">Height of the overlay tab bar in pixels.</param>
    public TabGroupController(IWindowManager windowManager, IntPtr window1, IntPtr window2, int tabBarHeight = 32)
    {
        _windowManager = windowManager ?? throw new ArgumentNullException(nameof(windowManager));

        if (window1 == IntPtr.Zero) throw new ArgumentException("Invalid window handle.", nameof(window1));
        if (window2 == IntPtr.Zero) throw new ArgumentException("Invalid window handle.", nameof(window2));

        GroupId = Guid.NewGuid();
        CreatedAt = DateTime.UtcNow;
        _overlayTracker = new OverlayTracker();

        // Build TabItem models from the window handles.
        var tab1 = CreateTabItemFromHandle(window1);
        var tab2 = CreateTabItemFromHandle(window2);

        var vm1 = new TabItemViewModel(tab1) { IsActive = true };
        var vm2 = new TabItemViewModel(tab2) { IsActive = false };

        Tabs.Add(vm1);
        Tabs.Add(vm2);
        ActiveTab = vm1;

        // Position window2 at window1's bounds, then hide it.
        var bounds = _windowManager.GetBounds(window1);
        _windowManager.SetBounds(window2, bounds.X, bounds.Y, bounds.Width, bounds.Height);
        _windowManager.Hide(window2);

        // Create and show the overlay.
        CreateOverlay(tabBarHeight);
    }

    // ─── Overlay Management ─────────────────────────────────────────────

    /// <summary>
    /// Creates the <see cref="TabOverlayWindow"/> and positions it above the active window.
    /// </summary>
    private void CreateOverlay(int tabBarHeight)
    {
        Overlay = new TabOverlayWindow(tabBarHeight);

        // Bind the overlay's tab collection to ours.
        foreach (var tab in Tabs)
            Overlay.Tabs.Add(tab);

        // Wire up overlay events.
        Overlay.TabClicked += OnOverlayTabClicked;
        Overlay.TabCloseRequested += OnOverlayTabCloseRequested;
        Overlay.TabMiddleClicked += OnOverlayTabMiddleClicked;
        Overlay.TabScrollRequested += OnOverlayTabScrollRequested;
        Overlay.AddTabRequested += OnOverlayAddTabRequested;

        // Show the overlay.
        Overlay.ShowOverlay();
        UpdateOverlayPosition();

        // Start tracking the active window's position.
        StartTrackingActiveWindow();
    }

    /// <summary>
    /// Updates the overlay position to sit directly above the active window.
    /// </summary>
    public void UpdateOverlayPosition()
    {
        if (Overlay is null || ActiveTab is null) return;

        var bounds = _windowManager.GetBounds(ActiveTab.Handle);
        if (bounds.Width <= 0 || bounds.Height <= 0) return;

        Overlay.UpdatePosition(bounds.X, bounds.Y, bounds.Width);
    }

    // ─── Tab Switching ──────────────────────────────────────────────────

    /// <summary>
    /// Switches to the tab at the specified index.
    /// </summary>
    /// <param name="index">Zero-based tab index.</param>
    /// <returns>True if the switch succeeded.</returns>
    public bool SwitchToTab(int index)
    {
        if (index < 0 || index >= Tabs.Count) return false;

        var newTab = Tabs[index];
        if (newTab == ActiveTab) return true;

        var previousTab = ActiveTab;
        if (previousTab is null) return false;

        // 1. Get current active window bounds so the new window occupies the same space.
        var bounds = _windowManager.GetBounds(previousTab.Handle);

        // 2. Hide the current active window.
        _windowManager.Hide(previousTab.Handle);

        // 3. Position the new window at the same bounds.
        _windowManager.SetBounds(newTab.Handle, bounds.X, bounds.Y, bounds.Width, bounds.Height);

        // 4. Show and bring the new window to front.
        _windowManager.Show(newTab.Handle);
        _windowManager.BringToFront(newTab.Handle);

        // 5. Update active state.
        previousTab.IsActive = false;
        newTab.IsActive = true;
        ActiveTab = newTab;

        // 6. Reposition overlay and restart tracking.
        StopTrackingActiveWindow();
        StartTrackingActiveWindow();
        UpdateOverlayPosition();

        ActiveTabChanged?.Invoke(this, newTab);
        return true;
    }

    // ─── Tab Management ─────────────────────────────────────────────────

    /// <summary>
    /// Adds a window to this tab group.
    /// </summary>
    /// <param name="handle">The window handle to add.</param>
    /// <returns>The created <see cref="TabItemViewModel"/>, or null if the window is invalid or already in the group.</returns>
    public TabItemViewModel? AddTab(IntPtr handle)
    {
        if (handle == IntPtr.Zero) return null;

        // Don't add duplicates.
        if (Tabs.Any(t => t.Handle == handle))
            return null;

        var tabItem = CreateTabItemFromHandle(handle);
        var vm = new TabItemViewModel(tabItem) { IsActive = false };

        // Position the new window at the active tab's bounds and hide it.
        if (ActiveTab is not null)
        {
            var bounds = _windowManager.GetBounds(ActiveTab.Handle);
            _windowManager.SetBounds(handle, bounds.X, bounds.Y, bounds.Width, bounds.Height);
        }
        _windowManager.Hide(handle);

        Tabs.Add(vm);
        Overlay?.Tabs.Add(vm);

        TabAdded?.Invoke(this, vm);
        return vm;
    }

    /// <summary>
    /// Removes a window from this tab group and restores it as a standalone window.
    /// </summary>
    /// <param name="handle">The window handle to remove.</param>
    /// <returns>True if the tab was found and removed.</returns>
    public bool RemoveTab(IntPtr handle)
    {
        var vm = Tabs.FirstOrDefault(t => t.Handle == handle);
        if (vm is null) return false;

        bool wasActive = vm.IsActive;
        Tabs.Remove(vm);
        Overlay?.Tabs.Remove(vm);

        // Restore the window: show it and position it.
        if (ActiveTab is not null)
        {
            var bounds = _windowManager.GetBounds(ActiveTab.Handle);
            _windowManager.SetBounds(handle, bounds.X, bounds.Y, bounds.Width, bounds.Height);
        }
        _windowManager.Show(handle);

        // If we removed the active tab, switch to the first remaining tab.
        if (wasActive && Tabs.Count > 0)
        {
            SwitchToTab(0);
        }

        TabRemoved?.Invoke(this, vm);

        // If only one tab remains, disband the group.
        if (Tabs.Count <= 1)
        {
            DisbandGroup();
        }

        return true;
    }

    /// <summary>
    /// Closes a window (sends WM_CLOSE) and removes it from the group.
    /// </summary>
    /// <param name="handle">The window handle to close.</param>
    /// <returns>True if the tab was found and the close was initiated.</returns>
    public bool CloseTab(IntPtr handle)
    {
        var vm = Tabs.FirstOrDefault(t => t.Handle == handle);
        if (vm is null) return false;

        bool wasActive = vm.IsActive;

        // If this is the active tab, switch to another tab first.
        if (wasActive && Tabs.Count > 1)
        {
            int currentIndex = Tabs.IndexOf(vm);
            int nextIndex = currentIndex > 0 ? currentIndex - 1 : 1;
            SwitchToTab(nextIndex);
        }

        // Remove from collections.
        Tabs.Remove(vm);
        Overlay?.Tabs.Remove(vm);

        // Send WM_CLOSE to the window.
        _windowManager.Close(handle);

        TabRemoved?.Invoke(this, vm);

        // If only one or zero tabs remain, disband.
        if (Tabs.Count <= 1)
        {
            DisbandGroup();
        }

        return true;
    }

    /// <summary>
    /// Moves a tab left or right within the current group order.
    /// </summary>
    /// <param name="handle">Window handle of the tab to move.</param>
    /// <param name="offset">Relative move offset (-1 for left, +1 for right).</param>
    /// <returns>True if the tab order changed.</returns>
    public bool MoveTab(IntPtr handle, int offset)
    {
        if (offset == 0)
            return false;

        var vm = Tabs.FirstOrDefault(t => t.Handle == handle);
        if (vm is null)
            return false;

        int oldIndex = Tabs.IndexOf(vm);
        if (oldIndex < 0)
            return false;

        int newIndex = Math.Clamp(oldIndex + offset, 0, Tabs.Count - 1);
        if (newIndex == oldIndex)
            return false;

        Tabs.Move(oldIndex, newIndex);
        Overlay?.Tabs.Move(oldIndex, newIndex);

        for (int i = 0; i < Tabs.Count; i++)
        {
            Tabs[i].Model.Order = i;
        }

        return true;
    }

    // ─── Group Lifecycle ────────────────────────────────────────────────

    /// <summary>
    /// Disbands the group: restores all windows to standalone state and closes the overlay.
    /// </summary>
    public void DisbandGroup()
    {
        StopTrackingActiveWindow();

        // Get the bounds from the active tab to position all restored windows.
        (int X, int Y, int Width, int Height) restoreBounds = (0, 0, 800, 600);
        if (ActiveTab is not null)
        {
            restoreBounds = _windowManager.GetBounds(ActiveTab.Handle);
        }

        // Restore all windows: show them and position them at the last known bounds.
        foreach (var tab in Tabs)
        {
            tab.IsActive = false;
            _windowManager.SetBounds(tab.Handle, restoreBounds.X, restoreBounds.Y,
                restoreBounds.Width, restoreBounds.Height);
            _windowManager.Show(tab.Handle);
        }

        // Close the overlay.
        if (Overlay is not null)
        {
            UnwireOverlayEvents();
            Overlay.Close();
            Overlay = null;
        }

        ActiveTab = null;
        Tabs.Clear();

        Disbanded?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Called by the overlay tracker when the active window has moved or resized.
    /// Updates the overlay position to follow the window.
    /// </summary>
    public void HandleActiveWindowMoved()
    {
        UpdateOverlayPosition();
    }

    // ─── Window Position Tracking ───────────────────────────────────────

    private void StartTrackingActiveWindow()
    {
        if (ActiveTab is null) return;

        _overlayTracker.PositionChanged += OnTrackedWindowPositionChanged;
        _overlayTracker.WindowMinimized += OnTrackedWindowMinimized;
        _overlayTracker.WindowRestored += OnTrackedWindowRestored;
        _overlayTracker.WindowActivated += OnTrackedWindowActivated;

        _overlayTracker.Track(ActiveTab.Handle);
    }

    private void StopTrackingActiveWindow()
    {
        _overlayTracker.PositionChanged -= OnTrackedWindowPositionChanged;
        _overlayTracker.WindowMinimized -= OnTrackedWindowMinimized;
        _overlayTracker.WindowRestored -= OnTrackedWindowRestored;
        _overlayTracker.WindowActivated -= OnTrackedWindowActivated;

        _overlayTracker.StopTracking();
    }

    private void OnTrackedWindowPositionChanged(int x, int y, int width, int height)
    {
        Overlay?.Dispatcher.Invoke(() =>
        {
            Overlay?.UpdatePosition(x, y, width);
        });
    }

    private void OnTrackedWindowMinimized()
    {
        Overlay?.Dispatcher.Invoke(() => Overlay?.HideOverlay());
    }

    private void OnTrackedWindowRestored()
    {
        Overlay?.Dispatcher.Invoke(() =>
        {
            Overlay?.ShowOverlay();
            UpdateOverlayPosition();
        });
    }

    private void OnTrackedWindowActivated()
    {
        Overlay?.Dispatcher.Invoke(() =>
        {
            Overlay?.ShowOverlay();
            UpdateOverlayPosition();
        });
    }

    // ─── Overlay Event Handlers ─────────────────────────────────────────

    private void OnOverlayTabClicked(object? sender, int index)
    {
        SwitchToTab(index);
    }

    private void OnOverlayTabCloseRequested(object? sender, IntPtr handle)
    {
        CloseTab(handle);
    }

    private void OnOverlayTabMiddleClicked(object? sender, IntPtr handle)
    {
        CloseTab(handle);
    }

    private void OnOverlayTabScrollRequested(object? sender, int direction)
    {
        if (ActiveTab is null || Tabs.Count <= 1) return;

        int currentIndex = Tabs.IndexOf(ActiveTab);
        int newIndex = currentIndex + direction;

        // Wrap around.
        if (newIndex < 0) newIndex = Tabs.Count - 1;
        else if (newIndex >= Tabs.Count) newIndex = 0;

        SwitchToTab(newIndex);
    }

    private void OnOverlayAddTabRequested(object? sender, EventArgs e)
    {
        // The "+" button's action is handled by the TabGroupManager,
        // which may show a picker or perform auto-detection.
        // This controller simply raises the event for external handling.
    }

    private void UnwireOverlayEvents()
    {
        if (Overlay is null) return;

        Overlay.TabClicked -= OnOverlayTabClicked;
        Overlay.TabCloseRequested -= OnOverlayTabCloseRequested;
        Overlay.TabMiddleClicked -= OnOverlayTabMiddleClicked;
        Overlay.TabScrollRequested -= OnOverlayTabScrollRequested;
        Overlay.AddTabRequested -= OnOverlayAddTabRequested;
    }

    // ─── Helpers ────────────────────────────────────────────────────────

    private TabItem CreateTabItemFromHandle(IntPtr handle)
    {
        var windowInfo = _windowManager.GetWindowInfo(handle);

        return new TabItem
        {
            Handle = handle,
            Title = windowInfo?.Title ?? $"Window {handle}",
            ProcessName = windowInfo?.ProcessName ?? string.Empty,
            IconData = windowInfo?.IconData,
            Order = Tabs.Count
        };
    }

    /// <summary>
    /// Creates a snapshot of this controller's state as a <see cref="TabGroup"/> model.
    /// </summary>
    public TabGroup ToSnapshot()
    {
        var group = new TabGroup
        {
            Id = GroupId,
            Name = Name,
            ActiveHandle = ActiveTab?.Handle ?? IntPtr.Zero,
        };

        foreach (var tab in Tabs)
            group.Tabs.Add(tab.Model);

        if (ActiveTab is not null)
        {
            var bounds = _windowManager.GetBounds(ActiveTab.Handle);
            group.Left = bounds.X;
            group.Top = bounds.Y;
            group.Width = bounds.Width;
            group.Height = bounds.Height;
        }

        return group;
    }

    // ─── IDisposable ────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        StopTrackingActiveWindow();
        _overlayTracker.Dispose();

        if (Overlay is not null)
        {
            UnwireOverlayEvents();
            Overlay.Close();
            Overlay = null;
        }
    }
}
