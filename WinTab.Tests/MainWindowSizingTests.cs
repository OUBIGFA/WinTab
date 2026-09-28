using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using WinTab.Helpers;
using WinTab.UI.Views;

internal static class MainWindowSizingTests
{
    private static readonly Size LargeWorkArea = new(1920, 1400);

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("main window sizing expands past 900 without scrollbars on first launch", () => OnSta(InitialContentFit));
        yield return ("main window sizing fits short content and retains the minimum height", () => OnSta(ShortContent));
        yield return ("main window sizing keeps the initial size inside a small work area", () => OnSta(SmallWorkArea));
        yield return ("main window sizing restores user sizes above and below the initial limit", () => OnSta(RestoreUserSize));
        yield return ("main window sizing fits a saved height to the screen without a permanent limit", () => OnSta(RestoreOnSmallerScreen));
        yield return ("main window sizing fits a saved width to a narrower screen without a permanent limit", () => OnSta(RestoreWideUserSize));
        yield return ("main window sizing allows native resizing in both dimensions and scrolling", () => OnSta(ResizeAndScroll));
    }

    private static void InitialContentFit() => WithWindow(window =>
    {
        MainWindow.ApplyInitialSize(window, null, LargeWorkArea);
        Check.Equal(960d, window.Width);
        Check.That(window.Height > 1050, "First launch must fit the whole page, including content beyond 900 DIPs.");
        AssertNoScrollbars(window);
        AssertFreelyResizable(window);
    }, 1050);

    private static void ShortContent()
    {
        foreach (var contentHeight in new[] { 400d, 700d })
        {
            WithWindow(window =>
            {
                MainWindow.ApplyInitialSize(window, null, LargeWorkArea);
                Check.Equal(Math.Max(contentHeight + 2, window.MinHeight), window.Height);
                AssertNoScrollbars(window);
                AssertFreelyResizable(window);
            }, contentHeight);
        }
    }

    private static void SmallWorkArea() => WithWindow(window =>
    {
        MainWindow.ApplyInitialSize(window, null, new Size(800, 700));
        Check.That(window.Width <= 800, "The first window must stay within the screen width.");
        Check.Equal(700d, window.Height);
        AssertNoScrollbars(window);
        AssertFreelyResizable(window);
    });

    private static void RestoreUserSize()
    {
        foreach (var saved in new[] { new Size(1100, 1200), new Size(800, 650) })
        {
            WithWindow(window =>
            {
                MainWindow.ApplyInitialSize(window, saved, LargeWorkArea);
                Check.Equal(saved, new Size(window.Width, window.Height), "A user size must not be capped at the first-launch height.");
                AssertFreelyResizable(window);
            });
        }
    }

    private static void RestoreOnSmallerScreen() => WithWindow(window =>
    {
        MainWindow.ApplyInitialSize(window, new Size(960, 1600), new Size(1920, 1100));
        Check.Equal(1100d, window.Height, "Only the available work area limits restoration of an oversized height.");
        AssertFreelyResizable(window);
    });

    private static void RestoreWideUserSize() => WithWindow(window =>
    {
        MainWindow.ApplyInitialSize(window, new Size(2400, 1000), LargeWorkArea);
        Check.Equal(LargeWorkArea.Width, window.Width, "A size saved on a wider monitor must fit the current work area.");
        Check.Equal(1000d, window.Height, "A saved height above the initial limit must remain unchanged when it fits.");
        AssertFreelyResizable(window);
    });

    private static void ResizeAndScroll() => WithWindow(window =>
    {
        MainWindow.ApplyInitialSize(window, null, LargeWorkArea);
        // Exercise a real WPF frame, off screen and without activation. No Explorer hooks or user input.
        window.Show();
        var scroll = (ScrollViewer)window.Content;
        Resize(window, 1100, 1000);
        var largeViewport = scroll.ViewportHeight;
        Check.That(scroll.ScrollableHeight > 0, "Content beyond the chosen window size remains scrollable.");
        scroll.ScrollToBottom();
        window.UpdateLayout();
        Check.That(scroll.VerticalOffset > 0, "The bottom of long content must remain reachable.");

        Resize(window, 800, 650);
        Check.That(scroll.ViewportHeight < largeViewport, "The content viewport must follow manual resizing.");
        Resize(window, 1100, 1000);
    });

    private static void Resize(Window window, double width, double height)
    {
        window.Width = width;
        window.Height = height;
        window.UpdateLayout();
        Check.That(Math.Abs(window.ActualWidth - width) < 2, $"Native width must follow resizing: {window.ActualWidth} vs {width}.");
        Check.That(Math.Abs(window.ActualHeight - height) < 2, $"Native height must follow resizing: {window.ActualHeight} vs {height}.");
    }

    private static void AssertFreelyResizable(Window window)
    {
        Check.Equal(ResizeMode.CanResize, window.ResizeMode);
        Check.Equal(SizeToContent.Manual, window.SizeToContent);
        Check.That(double.IsPositiveInfinity(window.MaxHeight), "The initial height limit must not become a permanent maximum.");
        Check.That(double.IsPositiveInfinity(window.MaxWidth), "The initial width must not become a permanent maximum.");
    }

    private static void AssertNoScrollbars(Window window)
    {
        // Exercise layout in an actual native window, without activating it or starting Explorer hooks.
        window.Show();
        window.UpdateLayout();
        var scroll = (ScrollViewer)window.Content;
        Check.Equal(0d, scroll.ScrollableHeight, "The entire page must be visible on first launch.");
        Check.Equal(0d, scroll.ScrollableWidth);
        Check.Equal(Visibility.Collapsed, scroll.ComputedVerticalScrollBarVisibility);
        Check.Equal(Visibility.Collapsed, scroll.ComputedHorizontalScrollBarVisibility);
    }

    private static void WithWindow(Action<Window> body, double contentHeight = 1600)
    {
        var window = new Window
        {
            Width = 960,
            MinWidth = 720,
            MinHeight = 560,
            ShowActivated = false,
            ShowInTaskbar = false,
            WindowStartupLocation = WindowStartupLocation.Manual,
            Left = -32000,
            Top = -32000,
            Content = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Content = new Border { Height = contentHeight }
            }
        };
        // Match MainWindow's client-area frame so the test measures the same available content area.
        System.Windows.Shell.WindowChrome.SetWindowChrome(window, new System.Windows.Shell.WindowChrome
        {
            CaptionHeight = 46,
            ResizeBorderThickness = new Thickness(6),
            GlassFrameThickness = new Thickness(-1),
            UseAeroCaptionButtons = false
        });
        try { body(window); }
        finally { window.Close(); }
    }

    private static async Task OnSta(Action body)
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(body, CancellationToken.None, TaskCreationOptions.None, scheduler);
    }
}
