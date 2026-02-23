using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Extensions.DependencyInjection;
using WinTab.App.Views.Pages;

namespace WinTab.App.Views;

public partial class MainWindow : System.Windows.Window
{
    public MainWindow(IServiceProvider serviceProvider)
    {
        InitializeComponent();

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

        ScrollViewer? sv = FindAncestorScrollViewer(source) ?? FindFirstScrollViewer(RootNavigation);
        if (sv is null || sv.ScrollableHeight <= 0)
            return;

        sv.ScrollToVerticalOffset(sv.VerticalOffset - e.Delta);
        e.Handled = true;
    }

    private static ScrollViewer? FindAncestorScrollViewer(DependencyObject current)
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

    private static ScrollViewer? FindFirstScrollViewer(DependencyObject root)
    {
        int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(root);
        for (int i = 0; i < count; i++)
        {
            DependencyObject child = System.Windows.Media.VisualTreeHelper.GetChild(root, i);
            if (child is ScrollViewer sv)
                return sv;
            ScrollViewer? nested = FindFirstScrollViewer(child);
            if (nested is not null)
                return nested;
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
        if (App.IsExplicitShutdownRequested)
        {
            base.OnClosing(e);
            return;
        }

        e.Cancel = true;
        Hide();
        ShowInTaskbar = false;
    }

    /// <summary>
    /// Shows the window and brings it to front. Called from tray icon.
    /// </summary>
    public void BringToForeground()
    {
        Show();
        WindowState = WindowState.Normal;
        ShowInTaskbar = true;
        Activate();
    }
}
