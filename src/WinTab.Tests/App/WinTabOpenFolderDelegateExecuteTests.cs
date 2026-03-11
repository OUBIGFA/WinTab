using System;
using System.Reflection;
using FluentAssertions;
using WinTab.ShellBridge;
using Xunit;

namespace WinTab.Tests.App;

public sealed class WinTabOpenFolderDelegateExecuteTests
{
    [Fact]
    public void Execute_WhenRecycleBinParameter_ShouldKeepPipeRouting()
    {
        Type delegateType = typeof(WinTabOpenFolderDelegateExecute);
        FieldInfo sendField = delegateType.GetField("SendOpenFolderRequest", BindingFlags.Static | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("SendOpenFolderRequest hook not found.");
        FieldInfo fallbackField = delegateType.GetField("OpenFallbackTarget", BindingFlags.Static | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("OpenFallbackTarget hook not found.");

        object? originalSend = sendField.GetValue(null);
        object? originalFallback = fallbackField.GetValue(null);
        int sendCalls = 0;
        int fallbackCalls = 0;

        try
        {
            sendField.SetValue(null, (Func<string, nint, bool, bool>)((_, _, _) =>
            {
                sendCalls++;
                return true;
            }));
            fallbackField.SetValue(null, (Action<string>)(target =>
            {
                fallbackCalls++;
                target.Should().Be("::{645FF040-5081-101B-9F08-00AA002F954E}");
            }));

            var command = new WinTabOpenFolderDelegateExecute();
            command.SetParameters("::{645FF040-5081-101B-9F08-00AA002F954E}");

            int hr = command.Execute();

            hr.Should().Be(0);
            sendCalls.Should().Be(1,
                "shell namespace targets should stay on the installed pipe path so the running app can reuse Explorer tabs");
            fallbackCalls.Should().Be(0);
        }
        finally
        {
            sendField.SetValue(null, originalSend);
            fallbackField.SetValue(null, originalFallback);
        }
    }

    [Fact]
    public void Execute_WhenPhysicalFolderParameter_ShouldKeepPipeRouting()
    {
        Type delegateType = typeof(WinTabOpenFolderDelegateExecute);
        FieldInfo sendField = delegateType.GetField("SendOpenFolderRequest", BindingFlags.Static | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("SendOpenFolderRequest hook not found.");
        FieldInfo fallbackField = delegateType.GetField("OpenFallbackTarget", BindingFlags.Static | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("OpenFallbackTarget hook not found.");

        object? originalSend = sendField.GetValue(null);
        object? originalFallback = fallbackField.GetValue(null);
        int sendCalls = 0;
        int fallbackCalls = 0;

        try
        {
            sendField.SetValue(null, (Func<string, nint, bool, bool>)((target, _, _) =>
            {
                sendCalls++;
                target.Should().Be(@"C:\Windows");
                return true;
            }));
            fallbackField.SetValue(null, (Action<string>)(_ => fallbackCalls++));

            var command = new WinTabOpenFolderDelegateExecute();
            command.SetParameters(@"C:\Windows");

            int hr = command.Execute();

            hr.Should().Be(0);
            sendCalls.Should().Be(1);
            fallbackCalls.Should().Be(0);
        }
        finally
        {
            sendField.SetValue(null, originalSend);
            fallbackField.SetValue(null, originalFallback);
        }
    }
}
