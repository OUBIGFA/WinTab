using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using WinTab.Core;
using WinTab.Platform.Win32;

namespace WinTab.App;

public partial class GroupHostWindow : Window, INotifyPropertyChanged
{
    private const string DefaultGroupName = "Default";
    private readonly DispatcherTimer _cleanupTimer = new();
    private IntPtr _hostPanelHandle;
    private HostTab? _selectedTab;

    public event PropertyChangedEventHandler? PropertyChanged;

    public string GroupName { get; private set; }

    public ObservableCollection<HostTab> Tabs { get; } = new();

    public HostTab? SelectedTab
    {
        get => _selectedTab;
        set
        {
            if (Equals(_selectedTab, value))
            {
                return;
            }

            _selectedTab = value;
            OnPropertyChanged(nameof(SelectedTab));
            UpdateVisibility();
        }
    }

    public bool HasAttached => Tabs.Count > 0;

    public GroupHostWindow()
        : this(DefaultGroupName)
    {
    }

    public GroupHostWindow(string groupName)
    {
        GroupName = NormalizeGroupName(groupName);

        InitializeComponent();
        DataContext = this;
        Loaded += OnLoaded;
        Closed += OnClosed;

        UpdateTitle();

        _cleanupTimer.Interval = TimeSpan.FromSeconds(2);
        _cleanupTimer.Tick += (_, __) => CleanupClosedTabs();
        _cleanupTimer.Start();

        // Handle auto-close empty groups policy
        if (AppSettings.CurrentInstance?.AutoCloseEmptyGroups == true)
        {
            Tabs.CollectionChanged += OnTabsChanged;
        }
    }

    public void RenameGroup(string groupName)
    {
        var normalized = NormalizeGroupName(groupName);
        if (string.Equals(GroupName, normalized, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        GroupName = normalized;
        UpdateTitle();
        OnPropertyChanged(nameof(GroupName));
    }

    public void ApplyWindowState(GroupWindowState state)
    {
        if (state is null)
        {
            return;
        }

        var safeBounds = ClampToVirtualScreen(new Rect(state.Left, state.Top, state.Width, state.Height));

        if (safeBounds.Width > 0 && safeBounds.Height > 0)
        {
            Width = safeBounds.Width;
            Height = safeBounds.Height;
        }

        Left = safeBounds.Left;
        Top = safeBounds.Top;

        WindowState = state.State switch
        {
            GroupWindowStateMode.Maximized => WindowState.Maximized,
            GroupWindowStateMode.Minimized => WindowState.Minimized,
            _ => WindowState.Normal
        };
    }

    private static Rect ClampToVirtualScreen(Rect bounds)
    {
        var screenBounds = new Rect(
            SystemParameters.VirtualScreenLeft,
            SystemParameters.VirtualScreenTop,
            SystemParameters.VirtualScreenWidth,
            SystemParameters.VirtualScreenHeight);

        if (screenBounds.Width <= 0 || screenBounds.Height <= 0)
        {
            return bounds;
        }

        var width = double.IsNaN(bounds.Width) || double.IsInfinity(bounds.Width) ? 0 : Math.Max(bounds.Width, 0);
        var height = double.IsNaN(bounds.Height) || double.IsInfinity(bounds.Height) ? 0 : Math.Max(bounds.Height, 0);
        var left = double.IsNaN(bounds.Left) || double.IsInfinity(bounds.Left) ? screenBounds.Left : bounds.Left;
        var top = double.IsNaN(bounds.Top) || double.IsInfinity(bounds.Top) ? screenBounds.Top : bounds.Top;

        if (width > screenBounds.Width)
        {
            width = screenBounds.Width;
        }

        if (height > screenBounds.Height)
        {
            height = screenBounds.Height;
        }

        var maxLeft = screenBounds.Right - Math.Max(width, 1);
        var maxTop = screenBounds.Bottom - Math.Max(height, 1);

        left = Math.Min(Math.Max(left, screenBounds.Left), maxLeft);
        top = Math.Min(Math.Max(top, screenBounds.Top), maxTop);

        return new Rect(left, top, width, height);
    }

    public GroupWindowState CaptureWindowState()
    {
        var bounds = WindowState == WindowState.Normal ? new Rect(Left, Top, Width, Height) : RestoreBounds;
        var stateMode = WindowState switch
        {
            WindowState.Maximized => GroupWindowStateMode.Maximized,
            WindowState.Minimized => GroupWindowStateMode.Minimized,
            _ => GroupWindowStateMode.Normal
        };

        return new GroupWindowState
        {
            GroupName = GroupName,
            Left = bounds.Left,
            Top = bounds.Top,
            Width = bounds.Width,
            Height = bounds.Height,
            State = stateMode
        };
    }

    private void UpdateTitle()
    {
        Title = string.Equals(GroupName, DefaultGroupName, StringComparison.OrdinalIgnoreCase)
            ? "WinTab Host"
            : $"WinTab Host - {GroupName}";
    }

    private static string NormalizeGroupName(string? groupName) =>
        string.IsNullOrWhiteSpace(groupName) ? DefaultGroupName : groupName.Trim();

    public void AttachWindow(WindowInfo window)
    {
        if (window.Handle == IntPtr.Zero)
        {
            return;
        }

        EnsureHostPanelHandle();

        var existing = Tabs.FirstOrDefault(tab => tab.Handle == window.Handle);
        if (existing is not null)
        {
            SelectedTab = existing;
            return;
        }

        var reparented = WindowReparenting.AttachToHost(window.Handle, _hostPanelHandle);
        var tabItem = new HostTab(window, reparented);
        Tabs.Add(tabItem);
        SelectedTab = tabItem;
        OnPropertyChanged(nameof(HasAttached));
    }

    public bool DetachSelected()
    {
        if (SelectedTab is null)
        {
            return false;
        }

        return DetachTab(SelectedTab);
    }

    private bool DetachTab(HostTab tab)
    {
        Tabs.Remove(tab);
        WindowReparenting.Detach(tab.Reparented);

        if (SelectedTab == tab)
        {
            SelectedTab = Tabs.LastOrDefault();
        }

        OnPropertyChanged(nameof(HasAttached));
        return true;
    }

    public void DetachAll()
    {
        foreach (var tab in Tabs.ToList())
        {
            WindowReparenting.Detach(tab.Reparented);
        }

        Tabs.Clear();
        SelectedTab = null;
        OnPropertyChanged(nameof(HasAttached));
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        EnsureHostPanelHandle();
        UpdateVisibility();
    }

    private void OnClosed(object? sender, EventArgs e)
    {
        _cleanupTimer.Stop();
        DetachAll();
    }

    private void OnSizeChanged(object sender, SizeChangedEventArgs e)
    {
        ResizeSelected();
    }

    private void OnCloseTabClicked(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button button && button.Tag is HostTab tab)
        {
            DetachTab(tab);
        }
    }

    private void CleanupClosedTabs()
    {
        if (Tabs.Count == 0)
        {
            return;
        }

        var removed = false;
        foreach (var tab in Tabs.ToList())
        {
            if (!WindowActions.IsAlive(tab.Handle))
            {
                DetachTab(tab);
                removed = true;
            }
        }

        if (removed)
        {
            UpdateVisibility();
        }
    }

    private void EnsureHostPanelHandle()
    {
        if (_hostPanelHandle != IntPtr.Zero)
        {
            return;
        }

        _hostPanelHandle = HostPanel.Handle;
    }

    private void UpdateVisibility()
    {
        foreach (var tab in Tabs)
        {
            if (SelectedTab is not null && tab.Handle == SelectedTab.Handle)
            {
                WindowActions.Show(tab.Handle);
                ResizeSelected();
            }
            else
            {
                WindowActions.Hide(tab.Handle);
            }
        }
    }

    private void ResizeSelected()
    {
        if (SelectedTab is null || _hostPanelHandle == IntPtr.Zero)
        {
            return;
        }

        var size = HostPanel.ClientSize;
        WindowReparenting.ResizeToHost(SelectedTab.Handle, size.Width, size.Height);
    }

    private void OnPropertyChanged(string propertyName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    private void OnTabsChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
    {
        // Auto-close empty groups if the policy is enabled
        if (AppSettings.CurrentInstance?.AutoCloseEmptyGroups == true && Tabs.Count == 0)
        {
            // Delay closing to avoid issues during tab removal
            Dispatcher.BeginInvoke(new Action(Close), System.Windows.Threading.DispatcherPriority.Background);
        }
    }
}
