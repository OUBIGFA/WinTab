using WinTab.Core.Enums;

namespace WinTab.Core.Models;

public sealed class AppSettings
{
    // Startup
    public bool StartMinimized { get; set; } = false;
    public bool RunAtStartup { get; set; }

    // Language
    public Language Language { get; set; } = Language.Chinese;

    // Appearance
    public ThemeMode Theme { get; set; } = ThemeMode.Light;

    // Behavior
    public bool EnableExplorerOpenVerbInterception { get; set; } = true;
    public bool PersistExplorerOpenVerbInterceptionAcrossExit { get; set; } = false;
    public bool OpenNewTabFromActiveTabPath { get; set; } = true;
    public bool OpenChildFolderInNewTabFromActiveTab { get; set; } = false;
    public bool CloseTabOnDoubleClick { get; set; } = false;

    // Schema version for migration
    public int SchemaVersion { get; set; } = 2;
}
