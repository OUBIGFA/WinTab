using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace WinTab.Helpers;

public sealed class StaTaskScheduler : TaskScheduler, IDisposable
{
    private const int SlowStartMs = 2_000;
    private readonly Action? _beforeReady;
    private readonly Thread _staThread;
    private readonly object _gate = new();
    private readonly ConcurrentDictionary<int, Task> _tasks = new();
    private readonly TaskCompletionSource<Dispatcher> _ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly Dispatcher _dispatcher;
    private bool _disposed;

    public StaTaskScheduler() : this(null) { }

    /// <param name="beforeReady">Runs on the new thread before its dispatcher exists; tests use it to delay or fail the start.</param>
    internal StaTaskScheduler(Action? beforeReady)
    {
        _beforeReady = beforeReady;
        _staThread = new Thread(Run) { IsBackground = true, Name = "WinTab shell worker" };
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();
        // Creating a dispatcher waits on nothing outside this process, so a slow start (a busy system at
        // sign-in) is only late, never stuck; a fixed deadline would fail WinTab's whole start instead.
        if (!((IAsyncResult)_ready.Task).AsyncWaitHandle.WaitOne(SlowStartMs))
            Trace.TraceWarning($"Shell worker has not started after {SlowStartMs} ms; still waiting.");
        _dispatcher = _ready.Task.GetAwaiter().GetResult();
    }

    public override int MaximumConcurrencyLevel => 1;
    internal bool IsCurrentThread => Thread.CurrentThread == _staThread;
    internal int PendingCount => _tasks.Count;

    private void Run()
    {
        Dispatcher dispatcher;
        try
        {
            _beforeReady?.Invoke();
            dispatcher = Dispatcher.CurrentDispatcher;
            SynchronizationContext.SetSynchronizationContext(new DispatcherSynchronizationContext(dispatcher));
        }
        catch (Exception exception)
        {
            // The constructor reports the real failure instead of waiting for a dispatcher that never comes.
            _ready.TrySetException(exception);
            return;
        }
        _ready.TrySetResult(dispatcher);
        Dispatcher.Run();
    }

    protected override void QueueTask(Task task)
    {
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_tasks.Count >= 256)
                throw new InvalidOperationException("The shell work queue is full.");
            _tasks[task.Id] = task;
            _dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
            {
                try
                {
                    TryExecuteTask(task);
                }
                finally
                {
                    _tasks.TryRemove(task.Id, out _);
                }
            }));
        }
    }

    protected override bool TryExecuteTaskInline(Task task, bool taskWasPreviouslyQueued) =>
        IsCurrentThread && !taskWasPreviouslyQueued && TryExecuteTask(task);

    protected override IEnumerable<Task> GetScheduledTasks() => _tasks.Values.ToArray();

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed)
                return;
            _disposed = true;
            _dispatcher.BeginInvokeShutdown(DispatcherPriority.Background);
        }
        if (!IsCurrentThread && !_staThread.Join(2_000))
            Trace.TraceError("Shell worker shutdown timed out; its current operation is still running.");
    }
}
