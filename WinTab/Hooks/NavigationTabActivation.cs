using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace WinTab.Hooks;

/// <summary>The tabs of a window and its active tab, as seen at one moment.</summary>
internal sealed record NavigationTabObservation(nint[] Handles, nint ActiveTab);

internal enum NavigationActivationResult { Activated, AlreadyActive, Cancelled, Ambiguous, TimedOut, SelectionRejected }

/// <summary>What selecting the new tab came to: done, not possible yet, or not possible at all.</summary>
internal enum NavigationSelectOutcome { Selected, NotReady, Rejected }

/// <summary>
/// Brings the tab Explorer opened for a middle click to the front. The click's outcome is the one tab
/// that appears in the window and was not there when the button went down; nothing is guessed from an
/// index or a title alone.
/// </summary>
internal static class NavigationTabActivation
{
    internal static bool TryGetOnlyAddition<T>(T[] before, T[] after, out T added) where T : notnull
    {
        added = default!;
        var oldSet = new HashSet<T>(before);
        var newSet = new HashSet<T>(after);
        if (oldSet.Count != before.Length || newSet.Count != after.Length ||
            newSet.Count != oldSet.Count + 1 || !oldSet.IsSubsetOf(newSet))
            return false;
        added = newSet.First(value => !oldSet.Contains(value));
        return true;
    }

    public static async Task<NavigationActivationResult> RunAsync(
        nint[] before, nint sourceTab,
        Func<bool> isCurrent, Func<NavigationTabObservation> observe,
        Func<nint, NavigationSelectOutcome> select, CancellationToken cancellationToken,
        int timeoutMs = 1_500, int pollMs = 20)
    {
        var deadline = Environment.TickCount64 + timeoutMs;
        nint candidate = 0;
        var confirmedAt = 0L;
        var selected = false;
        try
        {
            while (Environment.TickCount64 < deadline)
            {
                if (cancellationToken.IsCancellationRequested || !isCurrent())
                    return NavigationActivationResult.Cancelled;
                var current = observe();
                if (cancellationToken.IsCancellationRequested || !isCurrent())
                    return NavigationActivationResult.Cancelled;

                // Lost or extra tabs mean another operation is competing with this click.
                if (!before.All(current.Handles.Contains) || current.Handles.Length > before.Length + 1)
                    return NavigationActivationResult.Ambiguous;

                if (TryGetOnlyAddition(before, current.Handles, out var newTab))
                {
                    if (candidate != 0 && candidate != newTab)
                        return NavigationActivationResult.Ambiguous;
                    if (current.ActiveTab == newTab)
                        return selected ? NavigationActivationResult.Activated : NavigationActivationResult.AlreadyActive;
                    if (current.ActiveTab != sourceTab)
                    {
                        ExplorerDebugLog.Write($"Navigation active tab changed source={sourceTab} target={newTab} active={current.ActiveTab}");
                        return NavigationActivationResult.Cancelled;
                    }

                    if (candidate == 0)
                    {
                        candidate = newTab;
                        confirmedAt = Environment.TickCount64;
                    }
                    // Require a second coherent observation; a late competing creation must not be guessed.
                    else if (!selected && Environment.TickCount64 - confirmedAt >= 40)
                    {
                        if (cancellationToken.IsCancellationRequested || !isCurrent())
                            return NavigationActivationResult.Cancelled;
                        switch (select(newTab))
                        {
                            case NavigationSelectOutcome.Selected:
                                selected = true;
                                break;
                            case NavigationSelectOutcome.Rejected:
                                return NavigationActivationResult.SelectionRejected;
                        }
                    }
                }
                else if (selected || candidate != 0 || current.ActiveTab != sourceTab)
                {
                    return NavigationActivationResult.Ambiguous;
                }
                await Task.Delay(pollMs, cancellationToken).ConfigureAwait(false);
            }
            return NavigationActivationResult.TimedOut;
        }
        catch (OperationCanceledException)
        {
            return NavigationActivationResult.Cancelled;
        }
    }
}
