using System;
using System.Threading;
using WinTab.Helpers;

namespace WinTab.Hooks;

/// <summary>
/// Owns cancellation for one restore attempt. Foreground and navigation observations stop the
/// attempt after user interaction; this policy owns no Explorer COM objects or native hooks.
/// </summary>
internal sealed class ExplorerSessionRestoreAttempt : IDisposable
{
    private readonly CancellationTokenSource _cancellation;
    private readonly object _gate = new();
    private bool _sawForeground;
    private bool _watchNavigation;
    private bool _disposed;

    public ExplorerSessionRestoreAttempt(CancellationToken lifetime, bool inForeground, WindowIdentity initialTab)
    {
        _cancellation = CancellationTokenSource.CreateLinkedTokenSource(lifetime);
        Token = _cancellation.Token;
        _sawForeground = inForeground;
        InitialTab = initialTab;
    }

    public WindowIdentity InitialTab { get; }
    public CancellationToken Token { get; }

    public void ObserveForeground(bool isTarget)
    {
        lock (_gate)
        {
            if (_disposed) return;
            if (isTarget)
                _sawForeground = true;
            else if (_sawForeground && !_cancellation.IsCancellationRequested)
            {
                ExplorerDebugLog.Write("Session restore cancelled: foreground left the launch window");
                _cancellation.Cancel();
            }
        }
    }

    public void ArmNavigation()
    {
        lock (_gate)
            if (!_disposed) _watchNavigation = true;
    }

    public void InitialTabNavigated()
    {
        lock (_gate)
            if (!_disposed && _watchNavigation && !_cancellation.IsCancellationRequested)
            {
                ExplorerDebugLog.Write("Session restore cancelled: initial tab navigated");
                _cancellation.Cancel();
            }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed) return;
            _disposed = true;
            _cancellation.Dispose();
        }
    }
}
