using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using FluentAssertions;
using WinTab.App.Services;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using Xunit;

namespace WinTab.Tests.App;

public sealed class ExplorerOpenVerbStartupServiceTests : IDisposable
{
    private readonly string _tempDir;
    private readonly Logger _logger;

    public ExplorerOpenVerbStartupServiceTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "WinTabOpenVerbStartupTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
        _logger = new Logger(Path.Combine(_tempDir, "startup-tests.log"));
    }

    [Fact]
    public void Start_WhenAutoConvertEnabled_ShouldArmInterceptionBeforeReturning()
    {
        var interceptor = new FakeExplorerOpenVerbInterceptor(
            startupSelfCheckDelayMs: 400,
            enableOrRepairDelayMs: 400);

        var service = new ExplorerOpenVerbStartupService(
            interceptor,
            _logger,
            isWindows11: static () => true,
            resolveLaunchExecutablePath: static () => @"C:\Program Files\WinTab\WinTab.exe",
            isStableOpenVerbHandlerPath: static _ => true,
            runInBackground: work => Task.Run(work));

        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = true,
            RunAtStartup = true,
        };

        var sw = Stopwatch.StartNew();
        service.Start(settings);
        sw.ElapsedMilliseconds.Should().BeGreaterOrEqualTo(700,
            "startup must finish arming the session interception path before the app reports itself ready, otherwise the first folder open after launch can still flash a temporary Explorer window");

        interceptor.StartupSelfCheckCalls.Should().Be(1);
        interceptor.StartupSelfCheckArguments.Should().ContainSingle().Which.Should().BeTrue();
        interceptor.StartupSelfCheckPersistAcrossRebootArguments.Should().ContainSingle().Which.Should().BeFalse();
        interceptor.EnableOrRepairCalls.Should().Be(1);
        interceptor.EnableOrRepairPersistAcrossRebootArguments.Should().ContainSingle().Which.Should().BeFalse();
        interceptor.DisableAndRestoreCalls.Should().Be(0);
    }

    [Fact]
    public void Start_WhenAutoConvertDisabled_ShouldRunSynchronouslyToAvoidRace()
    {
        var interceptor = new FakeExplorerOpenVerbInterceptor(
            startupSelfCheckDelayMs: 250,
            enableOrRepairDelayMs: 0);

        var service = new ExplorerOpenVerbStartupService(
            interceptor,
            _logger,
            isWindows11: static () => true,
            resolveLaunchExecutablePath: static () => @"C:\Program Files\WinTab\WinTab.exe",
            isStableOpenVerbHandlerPath: static _ => true,
            runInBackground: work => Task.Run(work));

        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = false,
            RunAtStartup = false,
        };

        var sw = Stopwatch.StartNew();
        service.Start(settings);
        sw.ElapsedMilliseconds.Should().BeGreaterOrEqualTo(200,
            "disable path should complete synchronously so the shell is not left hijacked while a background task is pending");

        interceptor.StartupSelfCheckCalls.Should().Be(1);
        interceptor.StartupSelfCheckArguments.Should().ContainSingle().Which.Should().BeFalse();
        interceptor.StartupSelfCheckPersistAcrossRebootArguments.Should().ContainSingle().Which.Should().BeFalse();
        interceptor.DisableAndRestoreCalls.Should().Be(1);
        interceptor.DisableAndRestoreDeleteBackupArguments.Should().ContainSingle().Which.Should().BeFalse(
            "disabling interception while the app remains installed should preserve the original backup for later restart or uninstall");
        interceptor.EnableOrRepairCalls.Should().Be(0);
    }

    [Fact]
    public async Task WaitForStartupConfigurationToFinish_ShouldSynchronizeBackgroundOperationBeforeExitPath()
    {
        var interceptor = new FakeExplorerOpenVerbInterceptor(
            startupSelfCheckDelayMs: 200,
            enableOrRepairDelayMs: 200);

        var service = new ExplorerOpenVerbStartupService(
            interceptor,
            _logger,
            isWindows11: static () => true,
            resolveLaunchExecutablePath: static () => @"C:\Program Files\WinTab\WinTab.exe",
            isStableOpenVerbHandlerPath: static _ => true,
            runInBackground: work => Task.Run(work));

        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = true,
            RunAtStartup = true,
        };

        service.Start(settings);
        service.WaitForStartupConfigurationToFinish(TimeSpan.FromSeconds(2));

        await interceptor.WaitForCompletionAsync(TimeSpan.FromMilliseconds(100));
        interceptor.EnableOrRepairCalls.Should().Be(1);
        interceptor.EnableOrRepairPersistAcrossRebootArguments.Should().ContainSingle().Which.Should().BeFalse();
    }

    [Fact]
    public async Task Start_WhenAutoConvertEnabledButChildFolderNewTabDisabled_ShouldStillRepairInterception()
    {
        var interceptor = new FakeExplorerOpenVerbInterceptor();
        var service = new ExplorerOpenVerbStartupService(
            interceptor,
            _logger,
            isWindows11: static () => true,
            resolveLaunchExecutablePath: static () => @"C:\Program Files\WinTab\WinTab.exe",
            isStableOpenVerbHandlerPath: static _ => true,
            runInBackground: work => Task.Run(work));

        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = true,
            OpenChildFolderInNewTabFromActiveTab = false,
            RunAtStartup = false,
            PersistExplorerOpenVerbInterceptionAcrossExit = false,
        };

        service.Start(settings);
        await interceptor.WaitForCompletionAsync(TimeSpan.FromSeconds(2));

        interceptor.StartupSelfCheckArguments.Should().ContainSingle().Which.Should().BeTrue();
        interceptor.StartupSelfCheckPersistAcrossRebootArguments.Should().ContainSingle().Which.Should().BeFalse();
        interceptor.EnableOrRepairCalls.Should().Be(1);
        interceptor.DisableAndRestoreCalls.Should().Be(0);
    }

    [Fact]
    public async Task Start_WhenAutoConvertEnabledAndChildFolderNewTabEnabled_ShouldRepairInterception()
    {
        var interceptor = new FakeExplorerOpenVerbInterceptor();
        var service = new ExplorerOpenVerbStartupService(
            interceptor,
            _logger,
            isWindows11: static () => true,
            resolveLaunchExecutablePath: static () => @"C:\Program Files\WinTab\WinTab.exe",
            isStableOpenVerbHandlerPath: static _ => true,
            runInBackground: work => Task.Run(work));

        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = true,
            OpenChildFolderInNewTabFromActiveTab = true,
            RunAtStartup = false,
            PersistExplorerOpenVerbInterceptionAcrossExit = false,
        };

        service.Start(settings);
        await interceptor.WaitForCompletionAsync(TimeSpan.FromSeconds(2));

        interceptor.EnableOrRepairCalls.Should().Be(1);
        interceptor.EnableOrRepairPersistAcrossRebootArguments.Should().ContainSingle().Which.Should().BeFalse(
            "when child folders should open in a new tab, WinTab should still repair the Explorer interception path without leaving a reboot-persistent hijack behind");
        interceptor.DisableAndRestoreCalls.Should().Be(0);
    }

    [Fact]
    public async Task Start_WhenAutoConvertEnabled_ShouldRepairEvenWithoutRunAtStartup()
    {
        var interceptor = new FakeExplorerOpenVerbInterceptor();
        var service = new ExplorerOpenVerbStartupService(
            interceptor,
            _logger,
            isWindows11: static () => true,
            resolveLaunchExecutablePath: static () => @"C:\Program Files\WinTab\WinTab.exe",
            isStableOpenVerbHandlerPath: static _ => true,
            runInBackground: work => Task.Run(work));

        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = true,
            OpenChildFolderInNewTabFromActiveTab = true,
            RunAtStartup = false,
            PersistExplorerOpenVerbInterceptionAcrossExit = false,
        };

        service.Start(settings);
        await interceptor.WaitForCompletionAsync(TimeSpan.FromSeconds(2));

        interceptor.EnableOrRepairCalls.Should().Be(1);
        interceptor.EnableOrRepairPersistAcrossRebootArguments.Should().ContainSingle().Which.Should().BeFalse(
            "startup registration drift should not matter once Explorer capture is enabled, and the shell hook must still remain session-only");
    }

    [Fact]
    public void ReconfigureForCurrentSettings_WhenAutoConvertEnabled_ShouldApplySynchronously()
    {
        var interceptor = new FakeExplorerOpenVerbInterceptor(
            startupSelfCheckDelayMs: 200,
            enableOrRepairDelayMs: 200);

        var service = new ExplorerOpenVerbStartupService(
            interceptor,
            _logger,
            isWindows11: static () => true,
            resolveLaunchExecutablePath: static () => @"C:\Program Files\WinTab\WinTab.exe",
            isStableOpenVerbHandlerPath: static _ => true,
            runInBackground: work => Task.Run(work));

        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = true,
            RunAtStartup = false,
        };

        var sw = Stopwatch.StartNew();
        service.ReconfigureForCurrentSettings(settings);

        sw.ElapsedMilliseconds.Should().BeGreaterOrEqualTo(350,
            "when the user enables the feature at runtime, WinTab must finish arming the open-verb interception path before returning so the next folder open does not flash a temporary Explorer window");
        interceptor.StartupSelfCheckCalls.Should().Be(1);
        interceptor.EnableOrRepairCalls.Should().Be(1);
        interceptor.DisableAndRestoreCalls.Should().Be(0);
    }

    [Fact]
    public void ReconfigureForCurrentSettings_WhenAutoConvertDisabled_ShouldRemoveInterceptionSynchronously()
    {
        var interceptor = new FakeExplorerOpenVerbInterceptor(startupSelfCheckDelayMs: 250);

        var service = new ExplorerOpenVerbStartupService(
            interceptor,
            _logger,
            isWindows11: static () => true,
            resolveLaunchExecutablePath: static () => @"C:\Program Files\WinTab\WinTab.exe",
            isStableOpenVerbHandlerPath: static _ => true,
            runInBackground: work => Task.Run(work));

        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = false,
            RunAtStartup = false,
        };

        var sw = Stopwatch.StartNew();
        service.ReconfigureForCurrentSettings(settings);

        sw.ElapsedMilliseconds.Should().BeGreaterOrEqualTo(200,
            "when the user disables the feature at runtime, WinTab must remove the session interception before returning so Explorer immediately goes back to native behavior");
        interceptor.StartupSelfCheckCalls.Should().Be(1);
        interceptor.EnableOrRepairCalls.Should().Be(0);
        interceptor.DisableAndRestoreCalls.Should().Be(1);
    }

    [Fact]
    public async Task Start_WhenAutoConvertDisabled_ShouldNotRepairInterception()
    {
        var interceptor = new FakeExplorerOpenVerbInterceptor();
        var service = new ExplorerOpenVerbStartupService(
            interceptor,
            _logger,
            isWindows11: static () => true,
            resolveLaunchExecutablePath: static () => @"C:\Program Files\WinTab\WinTab.exe",
            isStableOpenVerbHandlerPath: static _ => true,
            runInBackground: work => Task.Run(work));

        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = false,
            OpenChildFolderInNewTabFromActiveTab = true,
            RunAtStartup = true,
            PersistExplorerOpenVerbInterceptionAcrossExit = false,
        };

        service.Start(settings);
        await interceptor.WaitForCompletionAsync(TimeSpan.FromSeconds(2));

        interceptor.EnableOrRepairCalls.Should().Be(0);
        interceptor.DisableAndRestoreCalls.Should().Be(1);
    }

    public void Dispose()
    {
        _logger.Dispose();

        try
        {
            Directory.Delete(_tempDir, recursive: true);
        }
        catch
        {
            // ignore
        }
    }

    private sealed class FakeExplorerOpenVerbInterceptor : IExplorerOpenVerbInterceptor
    {
        private readonly int _startupSelfCheckDelayMs;
        private readonly int _enableOrRepairDelayMs;
        private readonly TaskCompletionSource<bool> _completed =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        public FakeExplorerOpenVerbInterceptor(
            int startupSelfCheckDelayMs = 0,
            int enableOrRepairDelayMs = 0)
        {
            _startupSelfCheckDelayMs = startupSelfCheckDelayMs;
            _enableOrRepairDelayMs = enableOrRepairDelayMs;
        }

        public int StartupSelfCheckCalls { get; private set; }
        public int EnableOrRepairCalls { get; private set; }
        public int DisableAndRestoreCalls { get; private set; }
        public List<bool> StartupSelfCheckArguments { get; } = [];
        public List<bool> StartupSelfCheckPersistAcrossRebootArguments { get; } = [];
        public List<bool> EnableOrRepairPersistAcrossRebootArguments { get; } = [];
        public List<bool> DisableAndRestoreDeleteBackupArguments { get; } = [];

        public void StartupSelfCheck(bool settingEnabled, bool persistAcrossReboot)
        {
            StartupSelfCheckCalls++;
            StartupSelfCheckArguments.Add(settingEnabled);
            StartupSelfCheckPersistAcrossRebootArguments.Add(persistAcrossReboot);

            if (_startupSelfCheckDelayMs > 0)
                Task.Delay(_startupSelfCheckDelayMs).GetAwaiter().GetResult();
        }

        public void EnableOrRepair(bool persistAcrossReboot)
        {
            EnableOrRepairCalls++;
            EnableOrRepairPersistAcrossRebootArguments.Add(persistAcrossReboot);

            if (_enableOrRepairDelayMs > 0)
                Task.Delay(_enableOrRepairDelayMs).GetAwaiter().GetResult();

            _completed.TrySetResult(true);
        }

        public void DisableAndRestore(bool deleteBackup = true)
        {
            DisableAndRestoreCalls++;
            DisableAndRestoreDeleteBackupArguments.Add(deleteBackup);
            _completed.TrySetResult(true);
        }

        public async Task WaitForCompletionAsync(TimeSpan timeout)
        {
            Task completed = await Task.WhenAny(_completed.Task, Task.Delay(timeout));
            completed.Should().Be(_completed.Task, "startup background operation should complete in time");
            await _completed.Task;
        }
    }
}
