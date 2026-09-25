using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;

namespace WinTab.Hooks;

/// <summary>
/// Keeps wheel gestures in arrival order while a cold tab-row hit test runs off the input hook. Throttling
/// uses input time, not the time UIA finishes, so a slow first read neither loses nor replays a fast flick.
/// </summary>
internal sealed class ExplorerTabWheelSwitchController : IDisposable
{
    private const int GestureLifetimeMs = 1_000;
    private const int MaxPendingInputs = 32;
    private readonly Func<nint, CancellationToken, Task<Rectangle?>> _getTabRow;
    private readonly Func<nint, int, Func<bool>, Task<bool>> _switch;
    private readonly object _gate = new();
    private readonly WheelSwitchThrottle _throttle = new();
    private CancellationTokenSource _lifetime = new();
    private Task _tail = Task.CompletedTask;
    private nint _throttleWindow;
    private int _generation;
    private int _pending;
    private bool _enabled = true;
    private bool _disposed;

    public ExplorerTabWheelSwitchController(Func<nint, CancellationToken, Task<Rectangle?>> getTabRow,
        Func<nint, int, Func<bool>, Task<bool>> switchTab)
    {
        _getTabRow = getTabRow;
        _switch = switchTab;
    }

    public Task<bool> QueueAsync(nint window, Point point, int delta, long receivedAt,
        WheelSwitchSensitivity sensitivity, Func<bool> isCurrent)
    {
        lock (_gate)
        {
            if (!_enabled || _disposed || delta == 0 || _pending >= MaxPendingInputs)
                return Task.FromResult(false);
            var generation = _generation;
            var token = _lifetime.Token;
            bool IsCurrent() => !token.IsCancellationRequested && generation == Volatile.Read(ref _generation) &&
                Environment.TickCount64 - receivedAt <= GestureLifetimeMs && isCurrent();
            _pending++;
            var task = _tail.ContinueWith(_ => ProcessAsync(window, point, delta, receivedAt, sensitivity, IsCurrent, token),
                CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.Default).Unwrap();
            _tail = task;
            return task;
        }
    }

    private async Task<bool> ProcessAsync(nint window, Point point, int delta, long receivedAt,
        WheelSwitchSensitivity sensitivity, Func<bool> isCurrent, CancellationToken token)
    {
        try
        {
            if (!isCurrent()) return false;
            var row = await _getTabRow(window, token);
            if (!isCurrent() || row is not { } bounds || !bounds.Contains(point))
                return false;

            int step;
            lock (_gate)
            {
                if (!isCurrent()) return false;
                if (_throttleWindow != window)
                {
                    _throttleWindow = window;
                    _throttle.Reset();
                }
                step = _throttle.Accept(delta, receivedAt, sensitivity);
            }
            return step != 0 && await _switch(window, step, isCurrent);
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested)
        {
            return false;
        }
        finally
        {
            lock (_gate) _pending--;
        }
    }

    public void Start()
    {
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_enabled) return;
            _lifetime = new CancellationTokenSource();
            _enabled = true;
        }
    }

    public void Stop()
    {
        lock (_gate)
        {
            if (!_enabled) return;
            _enabled = false;
            _generation++;
            _lifetime.Cancel();
            _lifetime.Dispose();
            _throttleWindow = 0;
            _throttle.Reset();
        }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            Stop();
            _disposed = true;
        }
    }
}
