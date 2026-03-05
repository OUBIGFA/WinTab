using System;
using System.IO;
using FluentAssertions;
using WinTab.App.Services;
using Xunit;

namespace WinTab.Tests.App;

public sealed class ExplorerOpenVerbHandlerTests
{
    [Fact]
    public void IsStableOpenVerbHandlerPath_WhenPathEmpty_ShouldBeFalse()
    {
        bool stable = ExplorerOpenVerbHandler.IsStableOpenVerbHandlerPath(string.Empty);
        stable.Should().BeFalse();
    }

    [Fact]
    public void IsStableOpenVerbHandlerPath_WhenPathPointsToExistingExe_ShouldBeTrue()
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "WinTabOpenVerbTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        string exePath = Path.Combine(tempDir, "WinTab.exe");

        try
        {
            File.WriteAllText(exePath, "stub");

            bool stable = ExplorerOpenVerbHandler.IsStableOpenVerbHandlerPath(exePath);
            stable.Should().BeTrue();
        }
        finally
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void TryHandleOpenFolderInvocation_WhenFirstTwoPipeAttemptsFailThenSucceed_ShouldAvoidFallbackLaunch()
    {
        string tempDir = CreateTempDirectory();
        int sendAttempts = 0;
        int fallbackCalls = 0;

        var originalSend = ExplorerOpenVerbHandler.SendOpenFolderRequest;
        var originalDelay = ExplorerOpenVerbHandler.DelayBetweenRetries;
        var originalFallback = ExplorerOpenVerbHandler.OpenFolderFallback;

        try
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = (path, foreground) =>
            {
                sendAttempts++;
                return sendAttempts >= 3;
            };
            ExplorerOpenVerbHandler.DelayBetweenRetries = static _ => { };
            ExplorerOpenVerbHandler.OpenFolderFallback = (_, _) => fallbackCalls++;

            bool handled = ExplorerOpenVerbHandler.TryHandleOpenFolderInvocation(
                [RegistryOpenVerbInterceptor.HandlerArgument, tempDir],
                logger: null);

            handled.Should().BeTrue();
            sendAttempts.Should().Be(3);
            fallbackCalls.Should().Be(0);
        }
        finally
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = originalSend;
            ExplorerOpenVerbHandler.DelayBetweenRetries = originalDelay;
            ExplorerOpenVerbHandler.OpenFolderFallback = originalFallback;
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void TryHandleOpenFolderInvocation_WhenAllPipeAttemptsFail_ShouldFallbackExactlyOnce()
    {
        string tempDir = CreateTempDirectory();
        int sendAttempts = 0;
        int fallbackCalls = 0;

        var originalSend = ExplorerOpenVerbHandler.SendOpenFolderRequest;
        var originalDelay = ExplorerOpenVerbHandler.DelayBetweenRetries;
        var originalFallback = ExplorerOpenVerbHandler.OpenFolderFallback;

        try
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = (path, foreground) =>
            {
                sendAttempts++;
                return false;
            };
            ExplorerOpenVerbHandler.DelayBetweenRetries = static _ => { };
            ExplorerOpenVerbHandler.OpenFolderFallback = (_, _) => fallbackCalls++;

            bool handled = ExplorerOpenVerbHandler.TryHandleOpenFolderInvocation(
                [RegistryOpenVerbInterceptor.HandlerArgument, tempDir],
                logger: null);

            handled.Should().BeTrue();
            sendAttempts.Should().Be(3);
            fallbackCalls.Should().Be(1);
        }
        finally
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = originalSend;
            ExplorerOpenVerbHandler.DelayBetweenRetries = originalDelay;
            ExplorerOpenVerbHandler.OpenFolderFallback = originalFallback;
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void TryHandleOpenFolderInvocation_WhenShellNamespacePath_ShouldForwardRequest()
    {
        const string shellNamespace = "::{645FF040-5081-101B-9F08-00AA002F954E}";
        int sendAttempts = 0;

        var originalSend = ExplorerOpenVerbHandler.SendOpenFolderRequest;
        var originalDelay = ExplorerOpenVerbHandler.DelayBetweenRetries;
        var originalFallback = ExplorerOpenVerbHandler.OpenFolderFallback;

        try
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = (path, foreground) =>
            {
                sendAttempts++;
                path.Should().Be(shellNamespace);
                return true;
            };
            ExplorerOpenVerbHandler.DelayBetweenRetries = static _ => { };
            ExplorerOpenVerbHandler.OpenFolderFallback = (_, _) => throw new InvalidOperationException("Fallback should not run.");

            bool handled = ExplorerOpenVerbHandler.TryHandleOpenFolderInvocation(
                [RegistryOpenVerbInterceptor.HandlerArgument, shellNamespace],
                logger: null);

            handled.Should().BeTrue();
            sendAttempts.Should().Be(1);
        }
        finally
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = originalSend;
            ExplorerOpenVerbHandler.DelayBetweenRetries = originalDelay;
            ExplorerOpenVerbHandler.OpenFolderFallback = originalFallback;
        }
    }

    [Fact]
    public void TryHandleOpenFolderInvocation_WhenOpenChildFolderInNewTabDisabledAndExplorerForeground_ShouldForwardToMainInstanceNotLaunchNewWindow()
    {
        // Regression test: when OpenChildFolderInNewTabFromActiveTab=false and Explorer is foreground,
        // the handler must forward to the main instance so it navigates the current tab in-place.
        // Previously the handler called OpenFolderFallback (explorer.exe "path") which incorrectly
        // opened a new window instead of navigating the current tab.
        string tempDir = CreateTempDirectory();
        int sendAttempts = 0;
        int fallbackCalls = 0;

        var originalSend = ExplorerOpenVerbHandler.SendOpenFolderRequest;
        var originalDelay = ExplorerOpenVerbHandler.DelayBetweenRetries;
        var originalFallback = ExplorerOpenVerbHandler.OpenFolderFallback;

        try
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = (_, _) => { sendAttempts++; return true; };
            ExplorerOpenVerbHandler.DelayBetweenRetries = static _ => { };
            ExplorerOpenVerbHandler.OpenFolderFallback = (_, _) => fallbackCalls++;

            bool handled = ExplorerOpenVerbHandler.TryHandleOpenFolderInvocation(
                [RegistryOpenVerbInterceptor.HandlerArgument, tempDir],
                logger: null);

            handled.Should().BeTrue();
            sendAttempts.Should().Be(1, "must forward to main instance so it navigates the current tab without opening a new window");
            fallbackCalls.Should().Be(0, "must not call OpenFolderFallback which would launch a new explorer.exe window");
        }
        finally
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = originalSend;
            ExplorerOpenVerbHandler.DelayBetweenRetries = originalDelay;
            ExplorerOpenVerbHandler.OpenFolderFallback = originalFallback;
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void TryHandleOpenFolderInvocation_AlwaysForwardsToMainInstance_RegardlessOfForegroundState()
    {
        // The handler is unconditional: it always forwards to the main instance.
        // Navigation intent (new tab vs current tab) is decided by the main instance
        // based on its settings.
        string tempDir = CreateTempDirectory();
        int sendAttempts = 0;

        var originalSend = ExplorerOpenVerbHandler.SendOpenFolderRequest;
        var originalDelay = ExplorerOpenVerbHandler.DelayBetweenRetries;
        var originalFallback = ExplorerOpenVerbHandler.OpenFolderFallback;

        try
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = (_, _) => { sendAttempts++; return true; };
            ExplorerOpenVerbHandler.DelayBetweenRetries = static _ => { };
            ExplorerOpenVerbHandler.OpenFolderFallback = (_, _) => throw new InvalidOperationException("Fallback should not run when pipe succeeds.");

            bool handled = ExplorerOpenVerbHandler.TryHandleOpenFolderInvocation(
                [RegistryOpenVerbInterceptor.HandlerArgument, tempDir],
                logger: null);

            handled.Should().BeTrue();
            sendAttempts.Should().Be(1);
        }
        finally
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = originalSend;
            ExplorerOpenVerbHandler.DelayBetweenRetries = originalDelay;
            ExplorerOpenVerbHandler.OpenFolderFallback = originalFallback;
            Directory.Delete(tempDir, recursive: true);
        }
    }

    private static string CreateTempDirectory()
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "WinTabOpenVerbRoutingTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        return tempDir;
    }
}
