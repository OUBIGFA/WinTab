namespace WinTab.Core.Interfaces;

public interface IWindowEventSource : IDisposable
{
    event EventHandler<IntPtr>? WindowShown;
    event EventHandler<IntPtr>? WindowDestroyed;
    event EventHandler<IntPtr>? WindowLocationChanged;
    event EventHandler<IntPtr>? WindowMinimized;
    event EventHandler<IntPtr>? WindowRestored;
    event EventHandler<IntPtr>? WindowForegroundChanged;
    event EventHandler<IntPtr>? WindowMoveSizeStarted;
    event EventHandler<IntPtr>? WindowMoveSizeEnded;
}
