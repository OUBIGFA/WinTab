using System.Drawing;

namespace WinTab.Models;

public class HotKeyEventArgs(HotKeyProfile profile, nint foregroundWindow, Point? mousePosition = null)
{
    public HotKeyProfile Profile { get; } = profile;
    public nint ForegroundWindow { get; } = foregroundWindow;
    public Point? MousePosition { get; } = mousePosition;
}

