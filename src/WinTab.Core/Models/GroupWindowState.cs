using WinTab.Core.Enums;

namespace WinTab.Core.Models;

public sealed class GroupWindowState
{
    public string GroupName { get; set; } = "Default";
    public double Left { get; set; }
    public double Top { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public GroupWindowStateMode State { get; set; } = GroupWindowStateMode.Normal;
    public int ActiveTabIndex { get; set; }
    public List<GroupWindowTabState> Tabs { get; set; } = [];
}
