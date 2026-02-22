using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;

using CoreTabItem = WinTab.Core.Models.TabItem;

namespace WinTab.TabHost;

/// <summary>
/// Observable wrapper around <see cref="CoreTabItem"/> that adds UI-specific state.
/// </summary>
public sealed class TabItemViewModel : INotifyPropertyChanged
{
    private bool _isActive;

    public CoreTabItem Model { get; }
    public IntPtr Handle => Model.Handle;
    public string Title => Model.Title;
    public string ProcessName => Model.ProcessName;
    public byte[]? IconData => Model.IconData;

    public bool IsActive
    {
        get => _isActive;
        set
        {
            if (_isActive == value) return;
            _isActive = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsActive)));
        }
    }

    public TabItemViewModel(CoreTabItem model)
    {
        Model = model ?? throw new ArgumentNullException(nameof(model));
    }

    /// <summary>
    /// Refresh bindable properties from the underlying model (e.g., after title change).
    /// </summary>
    public void Refresh()
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Title)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IconData)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ProcessName)));
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}

/// <summary>
/// A transparent, topmost overlay window that renders the tab bar above a grouped window.
/// It applies WS_EX_TOOLWINDOW and WS_EX_NOACTIVATE so it never appears in Alt-Tab
/// and never steals focus from the underlying windows.
/// </summary>
public partial class TabOverlayWindow : Window
{
    // ─── P/Invoke (local declarations for window style manipulation) ─────

    private const int GWL_EXSTYLE = -20;
    private const long WS_EX_TOOLWINDOW = 0x00000080L;
    private const long WS_EX_NOACTIVATE = 0x08000000L;

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW", SetLastError = true)]
    private static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
    private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetWindowPos(
        IntPtr hWnd, IntPtr hWndInsertAfter,
        int X, int Y, int cx, int cy, uint uFlags);

    private static readonly IntPtr HWND_TOPMOST = new(-1);
    private const uint SWP_NOMOVE = 0x0002;
    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOACTIVATE = 0x0010;
    private const uint SWP_SHOWWINDOW = 0x0040;

    // ─── State ──────────────────────────────────────────────────────────

    private IntPtr _hwnd;
    private readonly int _tabBarHeight;

    // ─── Bindable Properties ────────────────────────────────────────────

    /// <summary>
    /// The collection of tab view models displayed in the tab bar.
    /// </summary>
    public ObservableCollection<TabItemViewModel> Tabs { get; } = [];

    /// <summary>
    /// Tab bar height for XAML binding.
    /// </summary>
    public int TabBarHeight => _tabBarHeight;

    // ─── Events ─────────────────────────────────────────────────────────

    /// <summary>Fired when the user scrolls the mouse wheel on the tab bar.</summary>
    public event EventHandler<int>? TabScrollRequested;

    /// <summary>Fired when the user middle-clicks a tab. Argument is the tab handle.</summary>
    public event EventHandler<IntPtr>? TabMiddleClicked;

    /// <summary>Fired when the user left-clicks a tab. Argument is the tab index.</summary>
    public event EventHandler<int>? TabClicked;

    /// <summary>Fired when the user clicks the close button on a tab. Argument is the tab handle.</summary>
    public event EventHandler<IntPtr>? TabCloseRequested;

    /// <summary>Fired when the user clicks the "+" button.</summary>
    public event EventHandler? AddTabRequested;

    // ─── Constructor ────────────────────────────────────────────────────

    public TabOverlayWindow(int tabBarHeight = 32)
    {
        _tabBarHeight = tabBarHeight;
        DataContext = this;
        InitializeComponent();

        Loaded += OnLoaded;
        PreviewMouseWheel += OnPreviewMouseWheel;
    }

    // ─── Lifecycle ──────────────────────────────────────────────────────

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        ApplyToolWindowStyles();
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        // Cache the HWND for later positioning calls.
        var helper = new WindowInteropHelper(this);
        _hwnd = helper.Handle;
    }

    // ─── Window Style Manipulation ──────────────────────────────────────

    /// <summary>
    /// Applies WS_EX_TOOLWINDOW (hides from Alt-Tab) and WS_EX_NOACTIVATE
    /// (clicking doesn't steal focus) to this overlay window.
    /// </summary>
    private void ApplyToolWindowStyles()
    {
        var helper = new WindowInteropHelper(this);
        _hwnd = helper.Handle;

        if (_hwnd == IntPtr.Zero) return;

        IntPtr exStyle = GetWindowLongPtr(_hwnd, GWL_EXSTYLE);
        long style = exStyle.ToInt64();
        style |= WS_EX_TOOLWINDOW;
        style |= WS_EX_NOACTIVATE;
        SetWindowLongPtr(_hwnd, GWL_EXSTYLE, new IntPtr(style));

        // Force the style change to take effect.
        SetWindowPos(_hwnd, HWND_TOPMOST, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    }

    // ─── Public API ─────────────────────────────────────────────────────

    /// <summary>
    /// Updates the overlay position and width to sit directly above the target window.
    /// </summary>
    /// <param name="x">Left edge in screen coordinates.</param>
    /// <param name="y">Top edge in screen coordinates (overlay is placed above this).</param>
    /// <param name="width">Width of the target window.</param>
    public void UpdatePosition(int x, int y, int width)
    {
        if (_hwnd == IntPtr.Zero) return;

        // Position the overlay directly above the target window's top edge.
        int overlayY = y - _tabBarHeight;

        Left = x;
        Top = overlayY;
        Width = width;
        Height = _tabBarHeight;

        // Ensure overlay stays topmost without stealing activation.
        SetWindowPos(_hwnd, HWND_TOPMOST, x, overlayY, width, _tabBarHeight,
            SWP_NOACTIVATE | SWP_SHOWWINDOW);
    }

    /// <summary>
    /// Shows the overlay window.
    /// </summary>
    public void ShowOverlay()
    {
        if (Visibility != Visibility.Visible)
            Show();

        if (_hwnd != IntPtr.Zero)
        {
            SetWindowPos(_hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
        }
    }

    /// <summary>
    /// Hides the overlay window.
    /// </summary>
    public void HideOverlay()
    {
        Hide();
    }

    // ─── Input Handlers ─────────────────────────────────────────────────

    private void OnPreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        // Delta > 0 = scroll up (previous tab), Delta < 0 = scroll down (next tab)
        int direction = e.Delta > 0 ? -1 : 1;
        TabScrollRequested?.Invoke(this, direction);
        e.Handled = true;
    }

    private void TabItem_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement element && element.DataContext is TabItemViewModel vm)
        {
            int index = Tabs.IndexOf(vm);
            if (index >= 0)
            {
                TabClicked?.Invoke(this, index);
            }
        }
    }

    private void TabItem_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton == MouseButton.Middle &&
            sender is FrameworkElement element &&
            element.DataContext is TabItemViewModel vm)
        {
            TabMiddleClicked?.Invoke(this, vm.Handle);
            e.Handled = true;
        }
    }

    private void CloseTabButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is IntPtr handle)
        {
            TabCloseRequested?.Invoke(this, handle);
        }
    }

    private void AddTabButton_Click(object sender, RoutedEventArgs e)
    {
        AddTabRequested?.Invoke(this, EventArgs.Empty);
    }
}
