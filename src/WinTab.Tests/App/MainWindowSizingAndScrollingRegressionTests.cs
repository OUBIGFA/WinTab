using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using FluentAssertions;
using WinTab.App.Views;
using Xunit;

namespace WinTab.Tests.App;

public sealed class MainWindowSizingAndScrollingRegressionTests
{
    [Fact]
    public void MainWindow_ShouldComputeGenerousLaunchSizeWithinTypicalDesktopWorkArea()
    {
        Size launchSize = MainWindow.CalculatePreferredLaunchSize(new Size(1600, 1000));

        launchSize.Width.Should().BeGreaterThanOrEqualTo(1280,
            "first launch should expose the full settings composition without manual resize on a typical desktop");
        launchSize.Height.Should().BeGreaterThanOrEqualTo(840,
            "first launch should expose the full settings composition without manual resize on a typical desktop");
        launchSize.Width.Should().BeLessThanOrEqualTo(1568);
        launchSize.Height.Should().BeLessThanOrEqualTo(952);
    }

    [Fact]
    public void ResolveMouseWheelTarget_ShouldPreferScrollableFallbackOverStaticScrollViewer()
    {
        (ScrollViewer? Resolved, ScrollViewer Expected) result = RunInSta(() =>
        {
            var root = new StackPanel();
            var headerScroll = new ScrollViewer { Content = new Border() };
            var contentScroll = new ScrollViewer { Content = new Border() };
            var source = new Border();

            root.Children.Add(headerScroll);
            root.Children.Add(contentScroll);
            root.Children.Add(source);

            return (
                MainWindow.ResolveMouseWheelTarget(
                    source,
                    root,
                    scrollViewer => ReferenceEquals(scrollViewer, contentScroll)),
                contentScroll);
        });

        result.Resolved.Should().NotBeNull();
        result.Resolved.Should().BeSameAs(result.Expected,
            "mouse wheel fallback should route to the actual scrollable content region instead of a decorative non-scrollable viewer");
    }

    [Fact]
    public void ResolveMouseWheelTarget_ShouldPreferNearestScrollableAncestor()
    {
        (ScrollViewer? Resolved, ScrollViewer Expected) result = RunInSta(() =>
        {
            var root = new StackPanel();
            var ancestorScroll = new ScrollViewer();
            var source = new Border();

            ancestorScroll.Content = source;
            root.Children.Add(ancestorScroll);

            return (
                MainWindow.ResolveMouseWheelTarget(
                    source,
                    root,
                    scrollViewer => ReferenceEquals(scrollViewer, ancestorScroll)),
                ancestorScroll);
        });

        result.Resolved.Should().NotBeNull();
        result.Resolved.Should().BeSameAs(result.Expected,
            "mouse wheel should stay within the user's current scroll context when an ancestor viewer can already scroll");
    }

    [Fact]
    public void MainWindowXaml_ShouldDefineComfortableMinimumWindowSize()
    {
        string windowPath = GetProjectFilePath("WinTab.App", "Views", "MainWindow.xaml");
        string xaml = File.ReadAllText(windowPath);

        xaml.Should().Contain("MinWidth=\"980\"",
            "the shell should keep enough width for the navigation rail and two-column settings layout");
        xaml.Should().Contain("MinHeight=\"700\"",
            "the shell should keep enough height for page headers and at least one full settings section");
    }

    private static T RunInSta<T>(Func<T> func)
    {
        Exception? exception = null;
        T? result = default;

        var thread = new Thread(() =>
        {
            try
            {
                result = func();
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (exception is not null)
        {
            throw exception;
        }

        return result!;
    }

    private static string GetProjectFilePath(string projectFolder, params string[] parts)
    {
        string[] allParts = ["src", projectFolder, .. parts];
        return TestRepoPaths.GetFile(allParts);
    }
}
