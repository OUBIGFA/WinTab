using System;
using System.IO;
using FluentAssertions;
using WinTab.App.ExplorerTabUtilityPort;
using WinTab.Diagnostics;
using Xunit;

namespace WinTab.Tests.App;

public sealed class ShellComNavigatorTests
{
    [Fact]
    public void TryNavigateComTab_WhenFileSystemPath_ShouldNavigateWithStringArgument()
    {
        using var logger = new Logger(Path.Combine(Path.GetTempPath(), $"wintab-shell-nav-{Guid.NewGuid():N}.log"));
        var navigator = new ShellComNavigator(logger);
        var fakeTab = new FakeComTab();
        const string path = @"C:\Windows";

        bool ok = navigator.TryNavigateComTab(fakeTab, path);

        ok.Should().BeTrue();
        fakeTab.NavigateCallCount.Should().Be(1);
        fakeTab.LastNavigateArgument.Should().BeOfType<string>();
        fakeTab.LastNavigateArgument.Should().Be(path);
    }

    [Fact]
    public void TryNavigateComTab_WhenShellNamespace_ShouldNavigateViaNamespaceObject()
    {
        using var logger = new Logger(Path.Combine(Path.GetTempPath(), $"wintab-shell-nav-{Guid.NewGuid():N}.log"));
        var navigator = new ShellComNavigator(logger);
        var fakeTab = new FakeComTab();
        const string recycleBinNamespace = "::{645FF040-5081-101B-9F08-00AA002F954E}";

        bool ok = navigator.TryNavigateComTab(fakeTab, recycleBinNamespace);

        ok.Should().BeTrue();
        fakeTab.NavigateCallCount.Should().Be(1);
        fakeTab.LastNavigateArgument.Should().NotBeOfType<string>(
            "virtual shell namespaces must navigate through Shell.NameSpace objects to avoid class-not-registered errors");
    }

    public sealed class FakeComTab
    {
        public int NavigateCallCount { get; private set; }
        public object? LastNavigateArgument { get; private set; }

        public void Navigate2(object target)
        {
            NavigateCallCount++;
            LastNavigateArgument = target;
        }
    }
}
