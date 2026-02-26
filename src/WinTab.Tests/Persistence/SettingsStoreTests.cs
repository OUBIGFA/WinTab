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
            settings.OpenNewTabFromActiveTabPath.Should().BeTrue();
            settings.OpenChildFolderInNewTabFromActiveTab.Should().BeFalse();
            settings.CloseTabOnDoubleClick.Should().BeFalse();
            settings.Theme.Should().Be(ThemeMode.Light);
            settings.SchemaVersion.Should().Be(2);
        }
        finally
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void Load_WhenSchemaV1_MigratesPersistFlagFromInterceptionToggle()
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "WinTabSettingsStoreTests", Guid.NewGuid().ToString("N"));
        string settingsPath = Path.Combine(tempDir, "settings.json");

        try
        {
            Directory.CreateDirectory(tempDir);
            File.WriteAllText(settingsPath, """
{
  "EnableExplorerOpenVerbInterception": false,
  "SchemaVersion": 1
}
""");

            var store = new SettingsStore(settingsPath);
            AppSettings settings = store.Load();

            settings.EnableExplorerOpenVerbInterception.Should().BeFalse();
            settings.SchemaVersion.Should().Be(2);
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
                ShowTrayIcon = false,
                EnableExplorerOpenVerbInterception = false,
                OpenNewTabFromActiveTabPath = false,
                OpenChildFolderInNewTabFromActiveTab = true,
                CloseTabOnDoubleClick = true,
                Theme = ThemeMode.Dark,
                Language = Language.English
            };

            store.Save(source);
            AppSettings restored = store.Load();

            restored.RunAtStartup.Should().BeTrue();
            restored.StartMinimized.Should().BeTrue();
            restored.ShowTrayIcon.Should().BeFalse();
            restored.EnableExplorerOpenVerbInterception.Should().BeFalse();
            restored.OpenNewTabFromActiveTabPath.Should().BeFalse();
            restored.OpenChildFolderInNewTabFromActiveTab.Should().BeTrue();
            restored.CloseTabOnDoubleClick.Should().BeTrue();
            restored.Theme.Should().Be(ThemeMode.Dark);
            restored.Language.Should().Be(Language.English);
            restored.SchemaVersion.Should().Be(2);
        }
        finally
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, recursive: true);
        }
    }
}
