using System;
using System.Reflection;
using FluentAssertions;
using WinTab.App.Services;
using Xunit;

namespace WinTab.Tests.App;

public sealed class ExplorerOpenRequestServerTests
{
    [Fact]
    public void TryParseOpenExRequest_ShouldRejectInvalidHwnd()
    {
        bool handled = InvokeTryParseOpenExRequest("OPEN_EX not_a_number C:\\Temp", out string? path, out IntPtr foreground, out string? invalidReason);

        handled.Should().BeFalse();
        path.Should().BeNull();
        foreground.Should().Be(IntPtr.Zero);
        invalidReason.Should().Be("invalid hwnd");
    }

    [Fact]
    public void TryParseOpenExRequest_ShouldRejectEmptyPath()
    {
        bool handled = InvokeTryParseOpenExRequest("OPEN_EX 12345   ", out string? path, out IntPtr foreground, out string? invalidReason);

        handled.Should().BeFalse();
        path.Should().BeNull();
        foreground.Should().Be(IntPtr.Zero);
        invalidReason.Should().Be("missing path");
    }

    [Fact]
    public void TryParseOpenRequest_ShouldRejectEmptyPath()
    {
        bool handled = InvokeTryParseOpenRequest("OPEN    ", out string? path);

        handled.Should().BeFalse();
        path.Should().BeNull();
    }


    private static bool InvokeTryParseOpenExRequest(string line, out string? path, out IntPtr foreground, out string? invalidReason)
    {
        MethodInfo method = typeof(ExplorerOpenRequestServer).GetMethod(
            "TryParseOpenExRequest",
            BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("TryParseOpenExRequest not found.");

        object?[] args = [line, null, IntPtr.Zero, null];
        bool handled = (bool)(method.Invoke(null, args) ?? false);
        path = args[1] as string;
        foreground = args[2] is IntPtr hwnd ? hwnd : IntPtr.Zero;
        invalidReason = args[3] as string;
        return handled;
    }

    private static bool InvokeTryParseOpenRequest(string line, out string? path)
    {
        MethodInfo method = typeof(ExplorerOpenRequestServer).GetMethod(
            "TryParseOpenRequest",
            BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("TryParseOpenRequest not found.");

        object?[] args = [line, null];
        bool handled = (bool)(method.Invoke(null, args) ?? false);
        path = args[1] as string;
        return handled;
    }

}
