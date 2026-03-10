using System;
using System.Reflection;
using FluentAssertions;
using WinTab.Platform.Win32;
using Xunit;

namespace WinTab.Tests.App;

public sealed class NativeShellLauncherTests
{
    [Fact]
    public void TryOpen_WhenRecycleBinShellNamespace_ShouldOpenViaParsedPidlAndReleaseIt()
    {
        Type launcherType = typeof(NativeShellLauncher);
        FieldInfo parseField = launcherType.GetField("ParseDisplayNameToPidl", BindingFlags.Static | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("ParseDisplayNameToPidl hook not found.");
        FieldInfo openField = launcherType.GetField("OpenFolderByPidl", BindingFlags.Static | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("OpenFolderByPidl hook not found.");
        FieldInfo releaseField = launcherType.GetField("ReleasePidl", BindingFlags.Static | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("ReleasePidl hook not found.");

        object? originalParse = parseField.GetValue(null);
        object? originalOpen = openField.GetValue(null);
        object? originalRelease = releaseField.GetValue(null);
        nint parsedPidl = new(0x1234);
        int parseCalls = 0;
        int openCalls = 0;
        int releaseCalls = 0;

        try
        {
            parseField.SetValue(null, (Func<string, nint>)(candidate =>
            {
                parseCalls++;
                candidate.Should().Be(ShellNamespacePath.RecycleBinShellAlias);
                return parsedPidl;
            }));
            openField.SetValue(null, (Func<nint, bool>)(pidl =>
            {
                openCalls++;
                pidl.Should().Be(parsedPidl);
                return true;
            }));
            releaseField.SetValue(null, (Action<nint>)(pidl =>
            {
                releaseCalls++;
                pidl.Should().Be(parsedPidl);
            }));

            bool opened = NativeShellLauncher.TryOpen(ShellNamespacePath.RecycleBinShellAlias);

            opened.Should().BeTrue();
            parseCalls.Should().Be(1);
            openCalls.Should().Be(1);
            releaseCalls.Should().Be(1,
                "native shell launch must release the parsed PIDL after opening");
        }
        finally
        {
            parseField.SetValue(null, originalParse);
            openField.SetValue(null, originalOpen);
            releaseField.SetValue(null, originalRelease);
        }
    }

    [Fact]
    public void TryOpen_WhenTargetIsPhysicalFolder_ShouldRejectWithoutParsing()
    {
        Type launcherType = typeof(NativeShellLauncher);
        FieldInfo parseField = launcherType.GetField("ParseDisplayNameToPidl", BindingFlags.Static | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("ParseDisplayNameToPidl hook not found.");

        object? originalParse = parseField.GetValue(null);
        int parseCalls = 0;

        try
        {
            parseField.SetValue(null, (Func<string, nint>)(_ =>
            {
                parseCalls++;
                return new nint(0x1234);
            }));

            bool opened = NativeShellLauncher.TryOpen(@"C:\Windows");

            opened.Should().BeFalse();
            parseCalls.Should().Be(0,
                "PIDL-based native shell launch should only run for targets explicitly classified as native shell namespaces");
        }
        finally
        {
            parseField.SetValue(null, originalParse);
        }
    }
}
