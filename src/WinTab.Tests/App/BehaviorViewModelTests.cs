using System;
using System.IO;
using FluentAssertions;
using WinTab.App.ExplorerTabUtilityPort;
using WinTab.App.ViewModels;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Persistence;
using Xunit;

namespace WinTab.Tests.App;

public sealed class BehaviorViewModelTests
{
    [Fact]
    public void IsOpenChildFolderInNewTabOptionEnabled_ShouldFollowAutoConvertToggle()
    {
        using var context = new TestContext(enableAutoConvert: false);

        context.ViewModel.IsOpenChildFolderInNewTabOptionEnabled.Should().BeFalse();

        context.ViewModel.EnableAutoConvertExplorerWindows = true;

        context.ViewModel.IsOpenChildFolderInNewTabOptionEnabled.Should().BeTrue();
    }

    [Fact]
    public void EnableAutoConvertExplorerWindows_WhenChanged_ShouldPersistAutoConvertState()
    {
        using var context = new TestContext(enableAutoConvert: false);

        context.ViewModel.EnableAutoConvertExplorerWindows = true;

        context.Settings.EnableAutoConvertExplorerWindows.Should().BeTrue();
        context.Settings.EnableExplorerOpenVerbInterception.Should().BeFalse();
        context.AutoConvertController.LastSetEnabledValue.Should().BeTrue();
        context.AutoConvertController.CallCount.Should().Be(1);
    }

    [Fact]
    public void EnableAutoConvertExplorerWindows_WhenDisabled_ShouldPersistAutoConvertState()
    {
        using var context = new TestContext(enableAutoConvert: true);

        context.ViewModel.EnableAutoConvertExplorerWindows = false;

        context.Settings.EnableAutoConvertExplorerWindows.Should().BeFalse();
        context.Settings.EnableExplorerOpenVerbInterception.Should().BeFalse();
        context.AutoConvertController.LastSetEnabledValue.Should().BeFalse();
        context.AutoConvertController.CallCount.Should().Be(1);
    }

    private sealed class TestContext : IDisposable
    {
        private readonly string _tempDir;

        public AppSettings Settings { get; }
        public Logger Logger { get; }
        public SettingsStore SettingsStore { get; }
        public ExplorerTabMouseHookService MouseHookService { get; }
        public FakeAutoConvertController AutoConvertController { get; }
        public BehaviorViewModel ViewModel { get; }

        public TestContext(bool enableAutoConvert)
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "WinTabBehaviorVmTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);

            Logger = new Logger(Path.Combine(_tempDir, "test.log"));
            SettingsStore = new SettingsStore(Path.Combine(_tempDir, "settings.json"), Logger);
            Settings = new AppSettings
            {
                EnableAutoConvertExplorerWindows = enableAutoConvert,
                EnableExplorerOpenVerbInterception = false,
                CloseTabOnDoubleClick = false
            };

            MouseHookService = new ExplorerTabMouseHookService(Settings, Logger);
            AutoConvertController = new FakeAutoConvertController();

            ViewModel = new BehaviorViewModel(
                Settings,
                SettingsStore,
                MouseHookService,
                AutoConvertController);
        }

        public void Dispose()
        {
            MouseHookService.Dispose();
            SettingsStore.Dispose();
            Logger.Dispose();

            if (Directory.Exists(_tempDir))
                Directory.Delete(_tempDir, recursive: true);
        }
    }

    public sealed class FakeAutoConvertController : IExplorerAutoConvertController
    {
        public bool IsAutoConvertEnabled { get; private set; }
        public bool? LastSetEnabledValue { get; private set; }
        public int CallCount { get; private set; }

        public void SetAutoConvertEnabled(bool enabled)
        {
            IsAutoConvertEnabled = enabled;
            LastSetEnabledValue = enabled;
            CallCount++;
        }
    }
}
