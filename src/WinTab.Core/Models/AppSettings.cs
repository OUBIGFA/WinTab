using WinTab.Core.Enums;

namespace WinTab.Core.Models;

public sealed class AppSettings
{
    // Startup
    public bool StartMinimized { get; set; } = false;
    public bool RunAtStartup { get; set; }
    public bool EnableTrayIcon { get; set; } = true;
    public bool MinimizeToTrayOnClose { get; set; } = true;

    // Language
    public Language Language { get; set; } = Language.Chinese;

    // Appearance
    public ThemeMode Theme { get; set; } = ThemeMode.Light;
    public TabStyle TabStyle { get; set; } = TabStyle.Modern;
    public bool UseRoundedCorners { get; set; } = true;
    public bool UseMicaEffect { get; set; } = true;
    public int TabBarHeight { get; set; } = 32;
    public double TabBarOpacity { get; set; } = 0.95;

    // Behavior
    public bool EnableExplorerOpenVerbInterception { get; set; } = true;
    public bool AutoApplyRules { get; set; } = false;
    public bool AutoCloseEmptyGroups { get; set; } = true;
    public bool EnableDragToGroup { get; set; } = false;
    public bool EnableMiddleClickClose { get; set; } = true;
    public bool EnableScrollWheelTabSwitch { get; set; } = true;
    public bool GroupSameProcessWindows { get; set; }
    public int DragGroupDelay { get; set; } = 300; // milliseconds

    // Auto-grouping rules
    public List<AutoGroupRule> AutoGroupRules { get; set; } = [];

    // Session persistence
    public List<GroupWindowState> SavedGroupStates { get; set; } = [];
    public bool RestoreSessionOnStartup { get; set; } = true;

    // Exclusions
    public List<string> ExcludedProcesses { get; set; } = [];

    // Schema version for migration
    public int SchemaVersion { get; set; } = 1;
}
