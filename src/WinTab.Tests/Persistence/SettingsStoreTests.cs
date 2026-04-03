using System;
using System.IO;
using System.Reflection;
using FluentAssertions;
using WinTab.Core.Enums;
using WinTab.Core.Models;
using WinTab.Diagnostics;
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
            settings.EnableExplorerOpenVerbInterception.Should().BeFalse();
            settings.EnableAutoConvertExplorerWindows.Should().BeFalse();
            settings.OpenNewTabFromActiveTabPath.Should().BeTrue();
            settings.OpenChildFolderInNewTabFromActiveTab.Should().BeFalse();
            settings.CloseTabOnDoubleClick.Should().BeFalse();
            settings.Theme.Should().Be(ThemeMode.Light);
            settings.SchemaVersion.Should().Be(3);
        }
        finally
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void Load_WhenSchemaV2AndAutoConvertEnabled_AlignsInterceptionFlag()
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "WinTabSettingsStoreTests", Guid.NewGuid().ToString("N"));
        string settingsPath = Path.Combine(tempDir, "settings.json");

        try
        {
            Directory.CreateDirectory(tempDir);
            File.WriteAllText(settingsPath, """
{
  "EnableExplorerOpenVerbInterception": false,
  "EnableAutoConvertExplorerWindows": true,
  "SchemaVersion": 2
}
""");

            var store = new SettingsStore(settingsPath);
            AppSettings settings = store.Load();

            settings.EnableExplorerOpenVerbInterception.Should().BeTrue();
            settings.EnableAutoConvertExplorerWindows.Should().BeTrue();
            settings.SchemaVersion.Should().Be(3);
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
            restored.SchemaVersion.Should().Be(3);
        }
        finally
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void Load_WhenReadThrowsUnauthorizedAccess_ShouldReturnDefaults()
    {
        ConstructorInfo? ctor = typeof(SettingsStore).GetConstructor(
            BindingFlags.Public | BindingFlags.Instance,
            binder: null,
            [typeof(string), typeof(Logger), typeof(Func<string, bool>), typeof(Func<string, string>)],
            modifiers: null);

        ctor.Should().NotBeNull("settings load should be testable against access-denied scenarios");

        Func<string, bool> fileExists = _ => true;
        Func<string, string> readAllText = _ => throw new UnauthorizedAccessException("denied");

        var store = (SettingsStore?)ctor?.Invoke(["ignored", null!, fileExists, readAllText]);
        store.Should().NotBeNull();

        AppSettings settings = store!.Load();
        settings.SchemaVersion.Should().Be(3);
        settings.EnableExplorerOpenVerbInterception.Should().BeFalse();
        settings.EnableAutoConvertExplorerWindows.Should().BeFalse();
    }
}
