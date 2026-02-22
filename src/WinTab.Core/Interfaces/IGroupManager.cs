using WinTab.Core.Models;

namespace WinTab.Core.Interfaces;

public interface IGroupManager
{
    TabGroup CreateGroup(IntPtr window1, IntPtr window2);
    TabGroup? AddToGroup(Guid groupId, IntPtr window);
    bool RemoveFromGroup(IntPtr window);
    TabGroup? GetGroupForWindow(IntPtr window);
    IReadOnlyList<TabGroup> GetAllGroups();
    bool SwitchTab(Guid groupId, int tabIndex);
    bool CloseTab(Guid groupId, IntPtr window);
    bool MoveTab(Guid groupId, IntPtr window, int offset);
    bool RenameGroup(Guid groupId, string name);
    bool DisbandGroup(Guid groupId);

    event EventHandler<TabGroup>? GroupCreated;
    event EventHandler<TabGroup>? GroupDisbanded;
    event EventHandler<TabGroup>? TabSwitched;
    event EventHandler<TabGroup>? TabAdded;
    event EventHandler<TabGroup>? TabRemoved;
}
