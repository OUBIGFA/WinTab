using System.IO;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public sealed class AppCompanionSafetyTests
{
    [Fact]
    public void AppSource_ShouldHandleCompanionModeInsteadOfImmediatelyExiting()
    {
        string sourcePath = TestRepoPaths.GetFile(["src", "WinTab.App", "App.xaml.cs"]);
        string source = File.ReadAllText(sourcePath);

        source.Should().Contain("--wintab-companion",
            "the app must still expose a dedicated companion mode for crash-safe cleanup");
        source.Should().Contain("RunCompanionCleanup(e.Args)",
            "a dedicated companion mode must execute a real cleanup path instead of exiting immediately");
        source.Should().Contain("--watch-parent",
            "the companion mode must accept the parent PID so it can wait for an unexpected WinTab exit before restoring Explorer");
    }
}
