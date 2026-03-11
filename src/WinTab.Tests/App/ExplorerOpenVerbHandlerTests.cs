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
            ExplorerOpenVerbHandler.SendOpenFolderRequest = (_, _) =>
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
            ExplorerOpenVerbHandler.SendOpenFolderRequest = (_, _) =>
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
    public void TryHandleOpenFolderInvocation_WhenRecycleBinNamespace_ShouldKeepPipeRouting()
    {
        const string shellNamespace = "::{645FF040-5081-101B-9F08-00AA002F954E}";
        int sendAttempts = 0;
        int fallbackCalls = 0;

        var originalSend = ExplorerOpenVerbHandler.SendOpenFolderRequest;
        var originalDelay = ExplorerOpenVerbHandler.DelayBetweenRetries;
        var originalFallback = ExplorerOpenVerbHandler.OpenFolderFallback;

        try
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = (path, _) =>
            {
                sendAttempts++;
                path.Should().Be(shellNamespace);
                return true;
            };
            ExplorerOpenVerbHandler.DelayBetweenRetries = static _ => { };
            ExplorerOpenVerbHandler.OpenFolderFallback = (path, _) =>
            {
                fallbackCalls++;
                path.Should().Be(shellNamespace);
            };

            bool handled = ExplorerOpenVerbHandler.TryHandleOpenFolderInvocation(
                [RegistryOpenVerbInterceptor.HandlerArgument, shellNamespace],
                logger: null);

            handled.Should().BeTrue();
            sendAttempts.Should().Be(1,
                "shell namespace targets should stay on the direct WinTab pipe path so the running app can reuse or activate Explorer tabs");
            fallbackCalls.Should().Be(0);
        }
        finally
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = originalSend;
            ExplorerOpenVerbHandler.DelayBetweenRetries = originalDelay;
            ExplorerOpenVerbHandler.OpenFolderFallback = originalFallback;
        }
    }

    [Fact]
    public void TryHandleOpenFolderInvocation_WhenDirectChildBrowse_ShouldStillUseDirectForwardingPath()
    {
        string tempDir = CreateTempDirectory();
        string parent = Path.Combine(tempDir, "parent");
        string child = Path.Combine(parent, "child");
        Directory.CreateDirectory(parent);
        Directory.CreateDirectory(child);

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
                [RegistryOpenVerbInterceptor.HandlerArgument, child],
                logger: null);

            handled.Should().BeTrue();
            sendAttempts.Should().Be(1,
                "direct-child browse must stay on direct pipe forwarding path to avoid open-window-convert-close flicker");
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
    public void TryHandleOpenFolderInvocation_WhenOpenChildFolderInNewTabEnabled_ShouldForwardToMainInstance()
    {
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
