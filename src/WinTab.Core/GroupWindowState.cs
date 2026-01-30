namespace WinTab.Core;

public enum GroupWindowStateMode
{
    Normal = 0,
    Minimized = 1,
    Maximized = 2
}

public sealed class GroupWindowState
{
    public string GroupName { get; set; } = "Default";
    public double Left { get; set; }
    public double Top { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public GroupWindowStateMode State { get; set; } = GroupWindowStateMode.Normal;
}
