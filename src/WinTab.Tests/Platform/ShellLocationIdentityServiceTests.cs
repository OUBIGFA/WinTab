using System;
using System.IO;
using FluentAssertions;
using WinTab.Platform.Win32;
using Xunit;

namespace WinTab.Tests.Platform;

public sealed class ShellLocationIdentityServiceTests
{
    private readonly ShellLocationIdentityService _service = new();

    [Fact]
    public void NormalizeLocation_ShouldCanonicalizeShellGuids()
    {
        const string guid = "{645FF040-5081-101B-9F08-00AA002F954E}";

        string fromBareGuid = _service.NormalizeLocation(guid);
        string fromShellPrefix = _service.NormalizeLocation("shell:::" + guid);
        string fromDoubleColon = _service.NormalizeLocation("::" + guid);

        fromBareGuid.Should().Be(fromShellPrefix);
        fromDoubleColon.Should().Be(fromShellPrefix);
        fromBareGuid.Should().StartWith("shell:::");
    }

    [Fact]
    public void AreEquivalent_ShouldMatchFilesystemPathAndFileUri()
    {
        string dir = Path.Combine(Path.GetTempPath(), "WinTabShellIdentityTests");
        Directory.CreateDirectory(dir);

        string fileUri = new Uri(dir + Path.DirectorySeparatorChar).AbsoluteUri;
        bool equivalent = _service.AreEquivalent(dir, fileUri);

        equivalent.Should().BeTrue();
    }

    [Fact]
    public void AreEquivalent_ShouldMatchDifferentPathFormatting()
    {
        string dir = Path.Combine(Path.GetTempPath(), "WinTabShellIdentityTests");
        Directory.CreateDirectory(dir);

        string withForwardSlash = dir.Replace('\\', '/');
        string withTrailingSlash = dir.EndsWith("\\", StringComparison.Ordinal) ? dir : dir + "\\";

        _service.AreEquivalent(withForwardSlash, withTrailingSlash).Should().BeTrue();
    }
}
