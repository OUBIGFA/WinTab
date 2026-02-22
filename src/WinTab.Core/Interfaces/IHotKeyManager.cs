using WinTab.Core.Enums;

namespace WinTab.Core.Interfaces;

public interface IHotKeyManager : IDisposable
{
    bool Register(int id, uint modifiers, uint key);
    bool Unregister(int id);
    void UnregisterAll();

    event EventHandler<HotKeyAction>? HotKeyPressed;
}
