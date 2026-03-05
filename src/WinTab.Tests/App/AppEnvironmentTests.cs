using System.IO;
using FluentAssertions;
using WinTab.App.Services;
using Xunit;

namespace WinTab.Tests.App;

public sealed class AppEnvironmentTests
{
    [Fact]
    public void TryNormalizeExistingDirectoryPath_WhenDriveDesignatorOnly_ShouldNormalizeToDriveRoot()
    {
        string originalCurrentDirectory = Directory.GetCurrentDirectory();
        string systemDirectory = Environment.SystemDirectory;
        string expectedDriveRoot = Path.GetPathRoot(systemDirectory)!;
        string driveDesignator = expectedDriveRoot[..2];

        try
        {
            Directory.SetCurrentDirectory(systemDirectory);

            bool ok = AppEnvironment.TryNormalizeExistingDirectoryPath(driveDesignator, out string normalized, out string reason);

            ok.Should().BeTrue();
            normalized.Should().Be(expectedDriveRoot);
            reason.Should().BeEmpty();
        }
        finally
        {
            Directory.SetCurrentDirectory(originalCurrentDirectory);
        }
    }

    [Fact]
    public void TryNormalizeExistingDirectoryPath_WhenShellNamespace_ShouldPassThrough()
    {
        const string shellPath = "::{645FF040-5081-101B-9F08-00AA002F954E}";

        bool ok = AppEnvironment.TryNormalizeExistingDirectoryPath(shellPath, out string normalized, out string reason);

        ok.Should().BeTrue();
        normalized.Should().Be(shellPath);
        reason.Should().BeEmpty();
    }

    [Fact]
    public void TryNormalizeExistingDirectoryPath_WhenUncPath_ShouldRemainAllowed()
    {
        const string uncPath = "\\\\server\\share";

        bool ok = AppEnvironment.TryNormalizeExistingDirectoryPath(uncPath, out string normalized, out string reason);

        ok.Should().BeTrue();
        normalized.Should().Be(uncPath);
        reason.Should().BeEmpty();
    }

    [Fact]
    public void TryNormalizeExistingDirectoryPath_WhenUriScheme_ShouldReject()
    {
        bool ok = AppEnvironment.TryNormalizeExistingDirectoryPath("file:///C:/Windows", out _, out string reason);

        ok.Should().BeFalse();
        reason.Should().Be("unsupported URI scheme");
    }

    [Fact]
    public void TryNormalizeExistingDirectoryPath_WhenBareGuid_ShouldNormalizeToShellNamespace()
    {
        const string bareGuid = "{645FF040-5081-101B-9F08-00AA002F954E}";

        bool ok = AppEnvironment.TryNormalizeExistingDirectoryPath(bareGuid, out string normalized, out string reason);

        ok.Should().BeTrue();
        Assert.Equal("::" + bareGuid, normalized);
        reason.Should().BeEmpty();
    }
}
