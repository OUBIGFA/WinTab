using WinTab.Core;
using WinTab.Platform.Win32;

namespace WinTab.App;

public sealed class HostTab
{
    public HostTab(WindowInfo window, ReparentedWindow reparented)
    {
        Window = window;
        Reparented = reparented;
    }

    public WindowInfo Window { get; }

    public ReparentedWindow Reparented { get; }

    public IntPtr Handle => Window.Handle;

    public string Title =>
        string.IsNullOrWhiteSpace(Window.Title) ? Window.ProcessName : Window.Title;
}
