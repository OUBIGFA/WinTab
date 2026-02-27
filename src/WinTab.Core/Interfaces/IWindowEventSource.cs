namespace WinTab.Core.Interfaces;

public interface IWindowEventSource : IDisposable
{
    event EventHandler<IntPtr>? WindowShown;
    event EventHandler<IntPtr>? WindowDestroyed;
    event EventHandler<IntPtr>? WindowForegroundChanged;
}
