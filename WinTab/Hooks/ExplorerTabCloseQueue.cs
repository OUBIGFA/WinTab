using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace WinTab.Hooks;

/// <summary>
/// Serializes close commands, not background observations. A delivered close releases the queue at once;
/// only a subsequent gesture that still sees the retiring tab waits for its replacement to become live.
/// </summary>
internal sealed class ExplorerTabCloseQueue(IExplorerTabCloseEnvironment environment, Func<int, Task>? delay = null)
{
    private readonly Func<int, Task> _delay = delay ?? (milliseconds => Task.Delay(milliseconds));
    private readonly object _gate = new();
    private Task _tail = Task.CompletedTask;
    private IssuedClose? _lastIssued;
    private int _pending;

    public bool HasPending => Volatile.Read(ref _pending) != 0;

    public Task<bool> QueueAsync(ExplorerTabCloseRequest request, Func<bool> isCurrent, Func<bool> canAct)
    {
        lock (_gate)
        {
            Interlocked.Increment(ref _pending);
            // Always continue, including after a provider exception. One failed close must not disable
            // later gestures; faults remain visible to the caller that submitted the failing request.
            var work = _tail.ContinueWith(_ => ExecuteAsync(request, isCurrent, canAct), CancellationToken.None,
                TaskContinuationOptions.None, TaskScheduler.Default).Unwrap();
            _tail = work;
            return work;
        }
    }

    private async Task<bool> ExecuteAsync(ExplorerTabCloseRequest request, Func<bool> isCurrent, Func<bool> canAct)
    {
        try
        {
            // Let the matching left-up drain before changing selection. This is per-command delivery,
            // not a debounce: the input hook has already rearmed for the next complete click pair.
            await _delay(10).ConfigureAwait(false);
            ExplorerTabAutomation.Tab[] tabs;
            while (true)
            {
                if (!isCurrent() || !canAct())
                    return false;
                tabs = environment.ReadTabs(request.ExplorerWindow);
                if (!isCurrent() || !canAct())
                    return false;

                var retiring = _lastIssued;
                var stillRetiring = retiring != null && retiring.Window == request.ExplorerWindow &&
                    retiring.IsCurrent() && tabs.Any(tab => tab.Id == retiring.TabId);
                if (tabs.Length != 0 && !stillRetiring)
                    break;

                // UIA can briefly expose an empty strip or the just-closed tab. Do not lose the next
                // gesture or invoke the old button twice. The request's ownership/deadline bounds retries.
                await _delay(15).ConfigureAwait(false);
            }

            var hit = tabs.Where(tab => tab.Bounds.Contains(request.Point.X, request.Point.Y)).ToArray();
            if (hit.Length != 1 || tabs.Any(tab => string.IsNullOrEmpty(tab.Id)) ||
                tabs.Select(tab => tab.Id).Distinct(StringComparer.Ordinal).Count() != tabs.Length)
            {
                ExplorerDebugLog.Write($"Double-click close has no unique live target tabs={tabs.Length} hits={hit.Length}");
                return false;
            }

            var closingId = hit[0].Id;
            var returnId = environment.GetReturnTab(request.ExplorerWindow, tabs, closingId);
            var delivered = environment.TryCloseTab(request.ExplorerWindow, closingId, returnId, canAct);
            if (delivered)
                _lastIssued = new IssuedClose(request.ExplorerWindow, closingId, canAct);
            return delivered;
        }
        finally
        {
            Interlocked.Decrement(ref _pending);
        }
    }

    private sealed record IssuedClose(nint Window, string TabId, Func<bool> IsCurrent);
}

/// <summary>Remote tab operations stay behind this boundary so queue timing can be replayed without desktop input.</summary>
internal interface IExplorerTabCloseEnvironment
{
    ExplorerTabAutomation.Tab[] ReadTabs(nint window);
    string? GetReturnTab(nint window, ExplorerTabAutomation.Tab[] tabs, string closingId);
    bool TryCloseTab(nint window, string closingId, string? returnId, Func<bool> isCurrent);
}
