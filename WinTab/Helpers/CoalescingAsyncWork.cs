using System;
using System.Diagnostics;
using System.Threading.Tasks;

namespace WinTab.Helpers;

internal sealed class CoalescingAsyncWork(Func<Task> action)
{
    private readonly object _gate = new();
    private TaskCompletionSource? _idle;
    private bool _pending;
    private bool _running;
    private bool _stopped;

    public Task WhenIdle
    {
        get { lock (_gate) return _idle?.Task ?? Task.CompletedTask; }
    }

    public void Request()
    {
        lock (_gate)
        {
            if (_stopped)
                return;
            _pending = true;
            if (_running)
                return;

            _running = true;
            _idle = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            _ = Task.Run(RunAsync);
        }
    }

    private async Task RunAsync()
    {
        while (true)
        {
            lock (_gate)
            {
                if (!_pending || _stopped)
                {
                    _running = false;
                    _idle!.TrySetResult();
                    return;
                }
                _pending = false;
            }

            try
            {
                await action().ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                Trace.TraceError($"Background work failed: {exception.GetType().Name}");
                lock (_gate)
                {
                    _running = false;
                    _pending = false;
                    _idle!.TrySetException(exception);
                }
                return;
            }
        }
    }

    public Task StopAsync()
    {
        lock (_gate)
        {
            _stopped = true;
            _pending = false;
            return _idle?.Task ?? Task.CompletedTask;
        }
    }
}
