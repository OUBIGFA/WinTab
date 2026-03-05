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
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = (_, _) => false;

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
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = (_, _) => false;

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
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = (_, _) => false;

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
    public void TryHandleOpenFolderInvocation_WhenOpenChildFolderInNewTabDisabled_ShouldStillForwardForDirectReuse()
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
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = (_, _) => false;

            bool handled = ExplorerOpenVerbHandler.TryHandleOpenFolderInvocation(
                [RegistryOpenVerbInterceptor.HandlerArgument, tempDir],
                logger: null);

            handled.Should().BeTrue();
            sendAttempts.Should().Be(1, "to avoid open-new-window-convert-close flicker, handler must keep direct forwarding reuse path");
            fallbackCalls.Should().Be(0, "direct reuse path must not launch native fallback when pipe path is available");
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
    public void TryHandleOpenFolderInvocation_WhenCurrentDirectoryBrowseAndChildFolderNewTabDisabled_ShouldUseNativeOpen()
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
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = (_, _) => true;

            bool handled = ExplorerOpenVerbHandler.TryHandleOpenFolderInvocation(
                [RegistryOpenVerbInterceptor.HandlerArgument, tempDir],
                logger: null);

            handled.Should().BeTrue();
            sendAttempts.Should().Be(0, "OFF mode current-directory browse should keep native behavior");
            fallbackCalls.Should().Be(1, "OFF mode current-directory browse should directly call native Explorer open");
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
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = (_, _) => false;

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

    [Fact]
    public void TryHandleOpenFolderInvocation_WhenOffButNotCurrentDirectoryBrowse_ShouldKeepForwardingPath()
    {
        string tempDir = CreateTempDirectory();
        int sendAttempts = 0;
        int fallbackCalls = 0;

        var originalSend = ExplorerOpenVerbHandler.SendOpenFolderRequest;
        var originalDelay = ExplorerOpenVerbHandler.DelayBetweenRetries;
        var originalFallback = ExplorerOpenVerbHandler.OpenFolderFallback;
        var originalBypass = ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen;
        var originalIsExplorer = ExplorerOpenVerbHandler.IsExplorerTopLevelWindowPredicate;
        var originalLoadSetting = ExplorerOpenVerbHandler.LoadOpenChildFolderInNewTabSettingPredicate;
        var originalGetDir = ExplorerOpenVerbHandler.TryGetForegroundExplorerDirectoryPredicate;

        try
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = (_, _) => { sendAttempts++; return true; };
            ExplorerOpenVerbHandler.DelayBetweenRetries = static _ => { };
            ExplorerOpenVerbHandler.OpenFolderFallback = (_, _) => fallbackCalls++;

            ExplorerOpenVerbHandler.IsExplorerTopLevelWindowPredicate = _ => true;
            ExplorerOpenVerbHandler.LoadOpenChildFolderInNewTabSettingPredicate = () => false;
            ExplorerOpenVerbHandler.TryGetForegroundExplorerDirectoryPredicate = _ => Path.Combine(Path.GetPathRoot(tempDir)!, "different-parent");
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen =
                (foreground, path) =>
                    ExplorerOpenVerbHandler.IsExplorerTopLevelWindowPredicate(foreground) &&
                    !ExplorerOpenVerbHandler.LoadOpenChildFolderInNewTabSettingPredicate() &&
                    string.Equals(ExplorerOpenVerbHandler.TryGetForegroundExplorerDirectoryPredicate(foreground), path, StringComparison.OrdinalIgnoreCase);

            bool handled = ExplorerOpenVerbHandler.TryHandleOpenFolderInvocation(
                [RegistryOpenVerbInterceptor.HandlerArgument, tempDir],
                logger: null);

            handled.Should().BeTrue();
            sendAttempts.Should().Be(1, "non current-directory browse in OFF mode should still use direct forwarding to avoid flash path");
            fallbackCalls.Should().Be(0);
        }
        finally
        {
            ExplorerOpenVerbHandler.SendOpenFolderRequest = originalSend;
            ExplorerOpenVerbHandler.DelayBetweenRetries = originalDelay;
            ExplorerOpenVerbHandler.OpenFolderFallback = originalFallback;
            ExplorerOpenVerbHandler.ShouldBypassInterceptionAndUseNativeOpen = originalBypass;
            ExplorerOpenVerbHandler.IsExplorerTopLevelWindowPredicate = originalIsExplorer;
            ExplorerOpenVerbHandler.LoadOpenChildFolderInNewTabSettingPredicate = originalLoadSetting;
            ExplorerOpenVerbHandler.TryGetForegroundExplorerDirectoryPredicate = originalGetDir;
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
