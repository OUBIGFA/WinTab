namespace WinTab.Core;

public sealed class AppSettings
{
    private static AppSettings? _currentInstance;
    
    public static AppSettings? CurrentInstance
    {
        get => _currentInstance;
        set => _currentInstance = value;
    }

    public bool StartMinimized { get; set; } = true;
    public bool EnableTrayIcon { get; set; } = true;
    public bool RunAtStartup { get; set; } = false;
    public bool AutoApplyRules { get; set; } = true;
    public bool AutoOpenGroups { get; set; } = false;

    public List<AutoGroupRule> AutoGroupRules { get; set; } = new();
    public List<GroupWindowState> GroupWindowStates { get; set; } = new();

    public List<WindowAttachment> WindowAttachments { get; set; } = new();

    public bool AutoCloseEmptyGroups { get; set; }

    public DateTime? LastSessionTime { get; set; }

    public Language Language { get; set; } = Language.Chinese;
}
