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
    public async Task Start_WhenInterceptorCallsAreSlow_ShouldNotBlockStartupThread()
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
            EnableExplorerOpenVerbInterception = true,
            OpenChildFolderInNewTabFromActiveTab = true,
        };

        var sw = Stopwatch.StartNew();
        service.Start(settings);
        sw.ElapsedMilliseconds.Should().BeLessThan(150,
            "startup should not wait on registry self-check/repair path");

        await interceptor.WaitForCompletionAsync(TimeSpan.FromSeconds(3));
        interceptor.StartupSelfCheckCalls.Should().Be(1);
        interceptor.EnableOrRepairCalls.Should().Be(1);
        interceptor.DisableAndRestoreCalls.Should().Be(0);
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
            EnableExplorerOpenVerbInterception = true,
            OpenChildFolderInNewTabFromActiveTab = true,
        };

        service.Start(settings);
        service.WaitForStartupConfigurationToFinish(TimeSpan.FromSeconds(2));

        await interceptor.WaitForCompletionAsync(TimeSpan.FromMilliseconds(100));
        interceptor.EnableOrRepairCalls.Should().Be(1);
    }
    [Fact]
    public async Task Start_WhenChildFolderNewTabDisabled_ShouldStillKeepInterceptionEnabledForDirectReuse()
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
            EnableExplorerOpenVerbInterception = true,
            OpenChildFolderInNewTabFromActiveTab = false,
        };

        service.Start(settings);
        await interceptor.WaitForCompletionAsync(TimeSpan.FromSeconds(2));

        interceptor.StartupSelfCheckArguments.Should().ContainSingle().Which.Should().BeTrue();
        interceptor.DisableAndRestoreCalls.Should().Be(0);
        interceptor.EnableOrRepairCalls.Should().Be(1);
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

        public void StartupSelfCheck(bool settingEnabled)
        {
            StartupSelfCheckCalls++;
            StartupSelfCheckArguments.Add(settingEnabled);

            if (_startupSelfCheckDelayMs > 0)
                Task.Delay(_startupSelfCheckDelayMs).GetAwaiter().GetResult();
        }

        public void EnableOrRepair()
        {
            EnableOrRepairCalls++;

            if (_enableOrRepairDelayMs > 0)
                Task.Delay(_enableOrRepairDelayMs).GetAwaiter().GetResult();

            _completed.TrySetResult(true);
        }

        public void DisableAndRestore()
        {
            DisableAndRestoreCalls++;
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


