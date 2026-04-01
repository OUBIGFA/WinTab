using System;
using System.Reflection;
using FluentAssertions;
using WinTab.App.ExplorerTabUtilityPort;
using Xunit;

namespace WinTab.Tests.App;

public sealed class ExplorerTabHookServiceConcurrencyTests
{
    [Fact]
    public void ClipboardGate_ShouldAllowOnlyOneConcurrentOwner()
    {
        MethodInfo tryEnter = typeof(ExplorerTabHookService).GetMethod(
            "TryEnterClipboardOperation",
            BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("TryEnterClipboardOperation method not found.");
        MethodInfo exit = typeof(ExplorerTabHookService).GetMethod(
            "ExitClipboardOperation",
            BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("ExitClipboardOperation method not found.");

        bool first = false;

        try
        {
            first = InvokeBool(tryEnter);
            bool second = InvokeBool(tryEnter);

            first.Should().BeTrue("the first clipboard owner should enter successfully");
            second.Should().BeFalse("a concurrent clipboard operation must be rejected to avoid clobbering the user's clipboard");

            exit.Invoke(null, null);
            first = false;

            bool third = InvokeBool(tryEnter);
            third.Should().BeTrue("releasing the gate should allow the next clipboard operation to proceed");
        }
        finally
        {
            if (first)
                exit.Invoke(null, null);
        }
    }

    private static bool InvokeBool(MethodInfo method)
    {
        return method.Invoke(null, null) as bool?
            ?? throw new InvalidOperationException($"Method did not return bool: {method.Name}");
    }
}
