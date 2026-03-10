using System.IO;
using System.Reflection;
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

    [Fact]
    public void TryOpenTargetFallback_WhenRecycleBinTarget_ShouldUseNativeShellLauncherInsteadOfExplorerProcess()
    {
        Type appEnvironmentType = typeof(AppEnvironment);
        FieldInfo nativeField = appEnvironmentType.GetField("TryOpenNativeShellTarget", BindingFlags.Static | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("TryOpenNativeShellTarget hook not found.");
        FieldInfo processField = appEnvironmentType.GetField("StartExplorerProcess", BindingFlags.Static | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("StartExplorerProcess hook not found.");

        object? originalNative = nativeField.GetValue(null);
        object? originalProcess = processField.GetValue(null);
        int nativeCalls = 0;
        int processCalls = 0;

        try
        {
            nativeField.SetValue(null, (Func<string, bool>)(target =>
            {
                nativeCalls++;
                target.Should().Be("::{645FF040-5081-101B-9F08-00AA002F954E}");
                return true;
            }));
            processField.SetValue(null, (Func<string, bool>)(_ =>
            {
                processCalls++;
                return true;
            }));

            bool opened = AppEnvironment.TryOpenTargetFallback("::{645FF040-5081-101B-9F08-00AA002F954E}", logger: null);

            opened.Should().BeTrue();
            nativeCalls.Should().Be(1);
            processCalls.Should().Be(0,
                "Recycle Bin fallback must stay on the native shell path and must not devolve to raw explorer.exe launching");
        }
        finally
        {
            nativeField.SetValue(null, originalNative);
            processField.SetValue(null, originalProcess);
        }
    }

    [Fact]
    public void TryOpenTargetFallback_WhenPhysicalFolderTarget_ShouldUseExplorerProcess()
    {
        Type appEnvironmentType = typeof(AppEnvironment);
        FieldInfo nativeField = appEnvironmentType.GetField("TryOpenNativeShellTarget", BindingFlags.Static | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("TryOpenNativeShellTarget hook not found.");
        FieldInfo processField = appEnvironmentType.GetField("StartExplorerProcess", BindingFlags.Static | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("StartExplorerProcess hook not found.");

        object? originalNative = nativeField.GetValue(null);
        object? originalProcess = processField.GetValue(null);
        int nativeCalls = 0;
        int processCalls = 0;

        try
        {
            nativeField.SetValue(null, (Func<string, bool>)(_ =>
            {
                nativeCalls++;
                return true;
            }));
            processField.SetValue(null, (Func<string, bool>)(target =>
            {
                processCalls++;
                target.Should().Be(@"C:\Windows");
                return true;
            }));

            bool opened = AppEnvironment.TryOpenTargetFallback(@"C:\Windows", logger: null);

            opened.Should().BeTrue();
            nativeCalls.Should().Be(0);
            processCalls.Should().Be(1);
        }
        finally
        {
            nativeField.SetValue(null, originalNative);
            processField.SetValue(null, originalProcess);
        }
    }
}
