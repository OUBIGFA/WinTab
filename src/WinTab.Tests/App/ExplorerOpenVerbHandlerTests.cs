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
        var originalBypass = ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen;

        try
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = (path, foreground) =>
            {
                sendAttempts++;
                return sendAttempts >= 3;
            };
            ExplorerOpenVerbHandler.DelayBetweenRetries = static _ => { };
            ExplorerOpenVerbHandler.OpenFolderFallback = (_, _) => fallbackCalls++;
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = static () => false;

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
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = originalBypass;
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
        var originalBypass = ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen;

        try
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = (path, foreground) =>
            {
                sendAttempts++;
                return false;
            };
            ExplorerOpenVerbHandler.DelayBetweenRetries = static _ => { };
            ExplorerOpenVerbHandler.OpenFolderFallback = (_, _) => fallbackCalls++;
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = static () => false;

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
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = originalBypass;
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
        var originalBypass = ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen;

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
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = static () => false;

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
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = originalBypass;
        }
    }

    [Fact]
    public void TryHandleOpenFolderInvocation_WhenOpenChildFolderInNewTabDisabled_ShouldUseNativeOpenDirectlyWithoutPipeForwarding()
    {
        string tempDir = CreateTempDirectory();
        int sendAttempts = 0;
        int fallbackCalls = 0;

        var originalSend = ExplorerOpenVerbHandler.SendOpenFolderRequest;
        var originalDelay = ExplorerOpenVerbHandler.DelayBetweenRetries;
        var originalFallback = ExplorerOpenVerbHandler.OpenFolderFallback;
        var originalBypass = ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen;

        try
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = (_, _) => { sendAttempts++; return true; };
            ExplorerOpenVerbHandler.DelayBetweenRetries = static _ => { };
            ExplorerOpenVerbHandler.OpenFolderFallback = (_, _) => fallbackCalls++;
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = static () => true;

            bool handled = ExplorerOpenVerbHandler.TryHandleOpenFolderInvocation(
                [RegistryOpenVerbInterceptor.HandlerArgument, tempDir],
                logger: null);

            handled.Should().BeTrue();
            sendAttempts.Should().Be(0, "setting disabled must bypass interception forwarding completely");
            fallbackCalls.Should().Be(1, "setting disabled must use native Explorer open directly");
        }
        finally
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = originalSend;
            ExplorerOpenVerbHandler.DelayBetweenRetries = originalDelay;
            ExplorerOpenVerbHandler.OpenFolderFallback = originalFallback;
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = originalBypass;
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void TryHandleOpenFolderInvocation_WhenOpenChildFolderInNewTabEnabled_ShouldForwardToMainInstance()
    {
        string tempDir = CreateTempDirectory();
        int sendAttempts = 0;

        var originalSend = ExplorerOpenVerbHandler.SendOpenFolderRequest;
        var originalDelay = ExplorerOpenVerbHandler.DelayBetweenRetries;
        var originalFallback = ExplorerOpenVerbHandler.OpenFolderFallback;
        var originalBypass = ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen;

        try
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = (_, _) => { sendAttempts++; return true; };
            ExplorerOpenVerbHandler.DelayBetweenRetries = static _ => { };
            ExplorerOpenVerbHandler.OpenFolderFallback = (_, _) => throw new InvalidOperationException("Fallback should not run when pipe succeeds.");
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = static () => false;

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
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = originalBypass;
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
