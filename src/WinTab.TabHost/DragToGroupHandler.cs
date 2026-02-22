using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using WinTab.Core.Interfaces;
using WinTab.Platform.Win32;

namespace WinTab.TabHost;

/// <summary>
/// Handles the drag-to-group gesture: when the user drags a window's title bar
/// over another window, a visual indicator appears and releasing the mouse
/// creates or extends a tab group.
/// </summary>
public sealed class DragToGroupHandler : IDisposable
{
    private readonly IGroupManager _groupManager;
    private readonly IWindowManager _windowManager;

    private DragDetector? _dragDetector;
    private bool _disposed;
    private bool _enabled;

    // Drag state.
    private IntPtr _sourceWindow;
    private string _sourceTitle = string.Empty;
    private IntPtr _currentTargetWindow;

    // Visual elements.
    private DragPreviewWindow? _dragPreview;
    private DropIndicatorWindow? _dropIndicator;

    // ─── P/Invoke (local declarations for WindowFromPoint) ──────────────

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    [DllImport("user32.dll")]
    private static extern IntPtr WindowFromPoint(POINT point);

    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    private const uint GA_ROOTOWNER = 3;
    private const int SM_CYCAPTION = 4;
    private const int SM_CYFRAME = 33;

    // ─── Constructor ────────────────────────────────────────────────────

    /// <summary>
    /// Creates a new <see cref="DragToGroupHandler"/>.
    /// </summary>
    /// <param name="groupManager">Group management service for creating/extending groups.</param>
    /// <param name="windowManager">Window information/manipulation service.</param>
    public DragToGroupHandler(IGroupManager groupManager, IWindowManager windowManager)
    {
        _groupManager = groupManager ?? throw new ArgumentNullException(nameof(groupManager));
        _windowManager = windowManager ?? throw new ArgumentNullException(nameof(windowManager));
    }

    // ─── Public API ─────────────────────────────────────────────────────

    /// <summary>
    /// Enables drag-to-group by subscribing to the specified <see cref="DragDetector"/>.
    /// </summary>
    /// <param name="dragDetector">The drag detector providing mouse gesture events.</param>
    public void Enable(DragDetector dragDetector)
    {
        if (_enabled || _disposed) return;

        _dragDetector = dragDetector ?? throw new ArgumentNullException(nameof(dragDetector));

        _dragDetector.DragStarted += OnDragStarted;
        _dragDetector.DragMoved += OnDragMoved;
        _dragDetector.DragEnded += OnDragEnded;
        _dragDetector.DragCancelled += OnDragCancelled;

        _enabled = true;
    }

    /// <summary>
    /// Disables drag-to-group and unsubscribes from the drag detector.
    /// </summary>
    public void Disable()
    {
        if (!_enabled || _dragDetector is null) return;

        _dragDetector.DragStarted -= OnDragStarted;
        _dragDetector.DragMoved -= OnDragMoved;
        _dragDetector.DragEnded -= OnDragEnded;
        _dragDetector.DragCancelled -= OnDragCancelled;

        CleanUpVisuals();
        _dragDetector = null;
        _enabled = false;
    }

    /// <summary>Whether drag-to-group is currently enabled.</summary>
    public bool IsEnabled => _enabled;

    // ─── Drag Event Handlers ────────────────────────────────────────────

    private void OnDragStarted(IntPtr sourceWindow, int x, int y)
    {
        _sourceWindow = sourceWindow;

        // Get source window title for the drag preview.
        var info = _windowManager.GetWindowInfo(sourceWindow);
        _sourceTitle = info?.Title ?? "Window";

        // Show drag preview near the cursor.
        Application.Current?.Dispatcher.Invoke(() =>
        {
            _dragPreview = new DragPreviewWindow(_sourceTitle);
            _dragPreview.UpdatePosition(x + 16, y + 16);
            _dragPreview.Show();
        });
    }

    private void OnDragMoved(int x, int y)
    {
        Application.Current?.Dispatcher.Invoke(() =>
        {
            // Update drag preview position.
            _dragPreview?.UpdatePosition(x + 16, y + 16);

            // Find the window under the cursor.
            var pt = new POINT { X = x, Y = y };
            IntPtr hwndUnderCursor = WindowFromPoint(pt);

            if (hwndUnderCursor == IntPtr.Zero)
            {
                HideDropIndicator();
                _currentTargetWindow = IntPtr.Zero;
                return;
            }

            // Walk up to the top-level (root owner) window.
            IntPtr rootOwner = GetAncestor(hwndUnderCursor, GA_ROOTOWNER);
            if (rootOwner != IntPtr.Zero)
                hwndUnderCursor = rootOwner;

            // Don't target the source window itself.
            if (hwndUnderCursor == _sourceWindow || hwndUnderCursor == IntPtr.Zero)
            {
                HideDropIndicator();
                _currentTargetWindow = IntPtr.Zero;
                return;
            }

            // Don't target overlay windows or the drag preview itself.
            if (IsOwnWindow(hwndUnderCursor))
            {
                HideDropIndicator();
                _currentTargetWindow = IntPtr.Zero;
                return;
            }

            // Check if the cursor is in the title bar region of the target window.
            if (IsInTitleBarRegion(hwndUnderCursor, x, y))
            {
                _currentTargetWindow = hwndUnderCursor;
                ShowDropIndicator(hwndUnderCursor);
            }
            else
            {
                HideDropIndicator();
                _currentTargetWindow = IntPtr.Zero;
            }
        });
    }

    private void OnDragEnded(int x, int y)
    {
        Application.Current?.Dispatcher.Invoke(() =>
        {
            try
            {
                if (_currentTargetWindow != IntPtr.Zero && _sourceWindow != IntPtr.Zero)
                {
                    PerformGroup(_sourceWindow, _currentTargetWindow);
                }
            }
            finally
            {
                CleanUpVisuals();
                _sourceWindow = IntPtr.Zero;
                _currentTargetWindow = IntPtr.Zero;
            }
        });
    }

    private void OnDragCancelled()
    {
        Application.Current?.Dispatcher.Invoke(() =>
        {
            CleanUpVisuals();
            _sourceWindow = IntPtr.Zero;
            _currentTargetWindow = IntPtr.Zero;
        });
    }

    // ─── Grouping Logic ─────────────────────────────────────────────────

    private void PerformGroup(IntPtr source, IntPtr target)
    {
        // Check if either window is already in a group.
        var sourceGroup = _groupManager.GetGroupForWindow(source);
        var targetGroup = _groupManager.GetGroupForWindow(target);

        if (sourceGroup is not null && targetGroup is not null)
        {
            // Both are already in groups. Don't merge groups automatically.
            return;
        }

        if (targetGroup is not null)
        {
            // Target is in a group; add source to that group.
            _groupManager.AddToGroup(targetGroup.Id, source);
        }
        else if (sourceGroup is not null)
        {
            // Source is in a group; add target to that group.
            _groupManager.AddToGroup(sourceGroup.Id, target);
        }
        else
        {
            // Neither is in a group; create a new group.
            _groupManager.CreateGroup(target, source);
        }
    }

    // ─── Visual Feedback ────────────────────────────────────────────────

    private void ShowDropIndicator(IntPtr targetWindow)
    {
        var bounds = _windowManager.GetBounds(targetWindow);
        if (bounds.Width <= 0 || bounds.Height <= 0) return;

        if (_dropIndicator is null)
        {
            _dropIndicator = new DropIndicatorWindow();
            _dropIndicator.Show();
        }

        _dropIndicator.UpdateBounds(bounds.X, bounds.Y, bounds.Width, bounds.Height);
    }

    private void HideDropIndicator()
    {
        _dropIndicator?.Hide();
    }

    private void CleanUpVisuals()
    {
        if (_dragPreview is not null)
        {
            _dragPreview.Close();
            _dragPreview = null;
        }

        if (_dropIndicator is not null)
        {
            _dropIndicator.Close();
            _dropIndicator = null;
        }
    }

    // ─── Helpers ────────────────────────────────────────────────────────

    /// <summary>
    /// Determines whether the screen point (x, y) is within the title bar region
    /// of the specified window.
    /// </summary>
    private static bool IsInTitleBarRegion(IntPtr hWnd, int screenX, int screenY)
    {
        if (!GetWindowRect(hWnd, out RECT rect))
            return false;

        int captionHeight = GetSystemMetrics(SM_CYCAPTION);
        int frameHeight = GetSystemMetrics(SM_CYFRAME);
        int titleBarBottom = rect.Top + frameHeight + captionHeight;

        return screenX >= rect.Left &&
               screenX <= rect.Right &&
               screenY >= rect.Top &&
               screenY <= titleBarBottom;
    }

    /// <summary>
    /// Checks whether the given handle belongs to one of our own WPF windows
    /// (drag preview, drop indicator, or overlay windows).
    /// </summary>
    private bool IsOwnWindow(IntPtr hwnd)
    {
        if (_dragPreview is not null)
        {
            var helper = new WindowInteropHelper(_dragPreview);
            if (helper.Handle == hwnd) return true;
        }

        if (_dropIndicator is not null)
        {
            var helper = new WindowInteropHelper(_dropIndicator);
            if (helper.Handle == hwnd) return true;
        }

        return false;
    }

    // ─── IDisposable ────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        Disable();
        CleanUpVisuals();
    }
}

// ═══════════════════════════════════════════════════════════════════════════
// DragPreviewWindow - small semi-transparent popup showing the dragged window's title
// ═══════════════════════════════════════════════════════════════════════════

/// <summary>
/// A small, semi-transparent WPF window that follows the cursor during a drag gesture,
/// showing the title of the window being dragged.
/// </summary>
internal sealed class DragPreviewWindow : Window
{
    // ─── P/Invoke for WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE ──────────────

    private const int GWL_EXSTYLE = -20;
    private const long WS_EX_TOOLWINDOW = 0x00000080L;
    private const long WS_EX_NOACTIVATE = 0x08000000L;
    private const long WS_EX_TRANSPARENT_STYLE = 0x00000020L;

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    private static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
    private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    public DragPreviewWindow(string title)
    {
        WindowStyle = WindowStyle.None;
        AllowsTransparency = true;
        Background = new SolidColorBrush(Color.FromArgb(200, 40, 40, 40));
        ShowInTaskbar = false;
        Topmost = true;
        ResizeMode = ResizeMode.NoResize;
        SizeToContent = SizeToContent.WidthAndHeight;
        IsHitTestVisible = false;

        var textBlock = new TextBlock
        {
            Text = title,
            Foreground = Brushes.White,
            FontSize = 12,
            Margin = new Thickness(8, 4, 8, 4),
            MaxWidth = 250,
            TextTrimming = TextTrimming.CharacterEllipsis,
            TextWrapping = TextWrapping.NoWrap
        };

        var border = new Border
        {
            CornerRadius = new CornerRadius(4),
            Background = new SolidColorBrush(Color.FromArgb(200, 40, 40, 40)),
            Child = textBlock,
            Padding = new Thickness(2)
        };

        Content = border;
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);

        var helper = new WindowInteropHelper(this);
        IntPtr hwnd = helper.Handle;

        if (hwnd != IntPtr.Zero)
        {
            IntPtr exStyle = GetWindowLongPtr(hwnd, GWL_EXSTYLE);
            long style = exStyle.ToInt64();
            style |= WS_EX_TOOLWINDOW;
            style |= WS_EX_NOACTIVATE;
            style |= WS_EX_TRANSPARENT_STYLE;
            SetWindowLongPtr(hwnd, GWL_EXSTYLE, new IntPtr(style));
        }
    }

    /// <summary>
    /// Moves the preview window to the specified screen coordinates.
    /// </summary>
    public void UpdatePosition(int x, int y)
    {
        Left = x;
        Top = y;
    }
}

// ═══════════════════════════════════════════════════════════════════════════
// DropIndicatorWindow - colored border overlay shown over valid drop targets
// ═══════════════════════════════════════════════════════════════════════════

/// <summary>
/// A transparent WPF window with a colored border that is shown over a valid
/// drop target window during a drag-to-group gesture. Acts as a visual hint
/// that releasing the mouse will create a tab group.
/// </summary>
internal sealed class DropIndicatorWindow : Window
{
    // ─── P/Invoke for WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE ──────────────

    private const int GWL_EXSTYLE = -20;
    private const long WS_EX_TOOLWINDOW = 0x00000080L;
    private const long WS_EX_NOACTIVATE = 0x08000000L;
    private const long WS_EX_TRANSPARENT_STYLE = 0x00000020L;

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    private static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
    private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    /// <summary>The accent color for the drop indicator border.</summary>
    private static readonly SolidColorBrush IndicatorBrush =
        new(Color.FromArgb(180, 0, 120, 215)); // Windows accent blue

    public DropIndicatorWindow()
    {
        WindowStyle = WindowStyle.None;
        AllowsTransparency = true;
        Background = Brushes.Transparent;
        ShowInTaskbar = false;
        Topmost = true;
        ResizeMode = ResizeMode.NoResize;
        IsHitTestVisible = false;

        var border = new Border
        {
            BorderBrush = IndicatorBrush,
            BorderThickness = new Thickness(3),
            CornerRadius = new CornerRadius(4),
            Background = new SolidColorBrush(Color.FromArgb(30, 0, 120, 215))
        };

        Content = border;
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);

        var helper = new WindowInteropHelper(this);
        IntPtr hwnd = helper.Handle;

        if (hwnd != IntPtr.Zero)
        {
            IntPtr exStyle = GetWindowLongPtr(hwnd, GWL_EXSTYLE);
            long style = exStyle.ToInt64();
            style |= WS_EX_TOOLWINDOW;
            style |= WS_EX_NOACTIVATE;
            style |= WS_EX_TRANSPARENT_STYLE;
            SetWindowLongPtr(hwnd, GWL_EXSTYLE, new IntPtr(style));
        }
    }

    /// <summary>
    /// Positions and sizes the indicator to cover the specified screen-space bounds.
    /// </summary>
    public void UpdateBounds(int x, int y, int width, int height)
    {
        Left = x;
        Top = y;
        Width = width;
        Height = height;

        if (Visibility != Visibility.Visible)
            Visibility = Visibility.Visible;
    }
}
