using WinTab.Core.Enums;

namespace WinTab.Core.Models;

public sealed class AutoGroupRule
{
    public AutoGroupMatchType MatchType { get; set; } = AutoGroupMatchType.ProcessName;
    public string MatchValue { get; set; } = string.Empty;
    public string GroupName { get; set; } = "Default";
    public int Priority { get; set; }
    public bool Enabled { get; set; } = true;
}
