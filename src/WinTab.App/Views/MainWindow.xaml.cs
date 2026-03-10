using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Extensions.DependencyInjection;
using WinTab.Core.Models;
using WinTab.App.Views.Pages;

namespace WinTab.App.Views;

public partial class MainWindow : Wpf.Ui.Controls.FluentWindow
{
    private readonly AppSettings _settings;

    public MainWindow(IServiceProvider serviceProvider, AppSettings settings)
    {
        InitializeComponent();
        _settings = settings;
        ApplyPreferredLaunchSize();

        AddHandler(UIElement.PreviewMouseWheelEvent, new MouseWheelEventHandler(OnAnyPreviewMouseWheel), handledEventsToo: true);

        // Set up page navigation service
        RootNavigation.SetServiceProvider(serviceProvider);

        Loaded += OnLoaded;
    }

    private void OnAnyPreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (e.Delta == 0)
            return;

        if (e.OriginalSource is not DependencyObject source)
            return;

        ScrollViewer? sv = ResolveMouseWheelTarget(source, RootNavigation);
        if (sv is null)
            return;

        double nextOffset = Math.Clamp(sv.VerticalOffset - e.Delta, 0, sv.ScrollableHeight);
        sv.ScrollToVerticalOffset(nextOffset);
        e.Handled = true;
    }

    internal static System.Windows.Size CalculatePreferredLaunchSize(System.Windows.Size workAreaSize)
    {
        double availableWidth = Math.Max(0, workAreaSize.Width - 32);
        double availableHeight = Math.Max(0, workAreaSize.Height - 48);

        double preferredWidth = Math.Clamp(workAreaSize.Width * 0.84, 1180, 1440);
        double preferredHeight = Math.Clamp(workAreaSize.Height * 0.88, 800, 940);

        return new System.Windows.Size(
            availableWidth <= 980 ? availableWidth : Math.Min(preferredWidth, availableWidth),
            availableHeight <= 700 ? availableHeight : Math.Min(preferredHeight, availableHeight));
    }

    internal static ScrollViewer? ResolveMouseWheelTarget(DependencyObject? source, DependencyObject? fallbackRoot)
    {
        return ResolveMouseWheelTarget(source, fallbackRoot, CanScroll);
    }

    internal static ScrollViewer? ResolveMouseWheelTarget(
        DependencyObject? source,
        DependencyObject? fallbackRoot,
        Func<ScrollViewer, bool> canScroll)
    {
        ArgumentNullException.ThrowIfNull(canScroll);

        ScrollViewer? ancestor = FindAncestorScrollViewer(source);
        if (ancestor is not null && canScroll(ancestor))
        {
            return ancestor;
        }

        return FindFirstScrollableDescendant(fallbackRoot, canScroll);
    }

    private void ApplyPreferredLaunchSize()
    {
        var workAreaSize = new System.Windows.Size(SystemParameters.WorkArea.Width, SystemParameters.WorkArea.Height);
        MinWidth = Math.Min(MinWidth, Math.Max(760, workAreaSize.Width - 32));
        MinHeight = Math.Min(MinHeight, Math.Max(560, workAreaSize.Height - 48));

        System.Windows.Size preferredSize = CalculatePreferredLaunchSize(workAreaSize);

        Width = Math.Max(preferredSize.Width, MinWidth);
        Height = Math.Max(preferredSize.Height, MinHeight);
    }

    private static bool CanScroll(ScrollViewer? scrollViewer)
    {
        return scrollViewer is not null
               && scrollViewer.ScrollableHeight > 0;
    }

    private static ScrollViewer? FindAncestorScrollViewer(DependencyObject? current)
    {
        DependencyObject? node = current;
        while (node is not null)
        {
            if (node is ScrollViewer sv)
                return sv;
            node = System.Windows.Media.VisualTreeHelper.GetParent(node);
        }
        return null;
    }

    private static ScrollViewer? FindFirstScrollableDescendant(
        DependencyObject? root,
        Func<ScrollViewer, bool> canScroll)
    {
        if (root is null)
        {
            return null;
        }

        ArgumentNullException.ThrowIfNull(canScroll);

        var queue = new Queue<DependencyObject>();
        queue.Enqueue(root);

        while (queue.Count > 0)
        {
            DependencyObject node = queue.Dequeue();

            if (node is ScrollViewer scrollViewer && canScroll(scrollViewer))
            {
                return scrollViewer;
            }

            int childrenCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(node);
            for (int i = 0; i < childrenCount; i++)
            {
                queue.Enqueue(System.Windows.Media.VisualTreeHelper.GetChild(node, i));
            }
        }

        return null;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        // Navigate to the General page by default
        RootNavigation.Navigate(typeof(GeneralPage));
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (!ShouldMinimizeToTrayOnClose(App.IsExplicitShutdownRequested, _settings.ShowTrayIcon))
        {
            base.OnClosing(e);
            return;
        }

        e.Cancel = true;
        Hide();
        ShowInTaskbar = false;
    }

    private static bool ShouldMinimizeToTrayOnClose(bool explicitShutdownRequested, bool trayIconVisible)
    {
        return !explicitShutdownRequested && trayIconVisible;
    }
}
