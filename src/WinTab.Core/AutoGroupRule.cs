namespace WinTab.Core;

public enum AutoGroupMatchType
{
    ProcessName = 0,
    WindowTitleContains = 1,
    ClassName = 2
}

public sealed class AutoGroupRule
{
    public string ProcessName { get; set; } = string.Empty;
    public AutoGroupMatchType MatchType { get; set; } = AutoGroupMatchType.ProcessName;
    public string MatchValue { get; set; } = string.Empty;
    public string GroupName { get; set; } = "Default";
    public int Priority { get; set; } = 0;
    public bool Enabled { get; set; } = true;
}
