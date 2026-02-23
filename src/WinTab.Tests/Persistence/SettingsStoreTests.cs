using System;
using System.IO;
using FluentAssertions;
using WinTab.Core.Enums;
using WinTab.Core.Models;
using WinTab.Persistence;
using Xunit;

namespace WinTab.Tests.Persistence;

public sealed class SettingsStoreTests
{
    [Fact]
    public void Load_WhenFileMissing_ReturnsDefaultSettings()
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "WinTabSettingsStoreTests", Guid.NewGuid().ToString("N"));
        string settingsPath = Path.Combine(tempDir, "settings.json");

        try
        {
            var store = new SettingsStore(settingsPath);

            AppSettings settings = store.Load();

            settings.RunAtStartup.Should().BeFalse();
            settings.EnableExplorerOpenVerbInterception.Should().BeTrue();
            settings.OpenChildFolderInNewTabFromActiveTab.Should().BeFalse();
            settings.Theme.Should().Be(ThemeMode.Light);
            settings.SchemaVersion.Should().Be(1);
        }
        finally
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void Save_ThenLoad_PreservesCurrentSettingsShape()
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "WinTabSettingsStoreTests", Guid.NewGuid().ToString("N"));
        string settingsPath = Path.Combine(tempDir, "settings.json");

        try
        {
            var store = new SettingsStore(settingsPath);
            var source = new AppSettings
            {
                RunAtStartup = true,
                StartMinimized = true,
                EnableExplorerOpenVerbInterception = false,
                OpenChildFolderInNewTabFromActiveTab = true,
                Theme = ThemeMode.Dark,
                Language = Language.English
            };

            store.Save(source);
            AppSettings restored = store.Load();

            restored.RunAtStartup.Should().BeTrue();
            restored.StartMinimized.Should().BeTrue();
            restored.EnableExplorerOpenVerbInterception.Should().BeFalse();
            restored.OpenChildFolderInNewTabFromActiveTab.Should().BeTrue();
            restored.Theme.Should().Be(ThemeMode.Dark);
            restored.Language.Should().Be(Language.English);
            restored.SchemaVersion.Should().Be(1);
        }
        finally
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, recursive: true);
        }
    }
}
