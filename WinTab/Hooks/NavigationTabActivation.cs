using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace WinTab.Hooks;

internal sealed record NavigationTabSet(nint[] Handles, string[] AutomationIds);
internal sealed record NavigationTabObservation(NavigationTabSet Tabs, nint ActiveTab);
internal enum NavigationActivationResult { Activated, AlreadyActive, Cancelled, Ambiguous, TimedOut, SelectionRejected }

/// <summary>Correlates both native and UIA identities. Never guesses an index or a tab title.</summary>
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
        NavigationTabSet before, nint sourceTab,
        Func<bool> isCurrent, Func<NavigationTabObservation> observe,
        Func<string, bool> select, CancellationToken cancellationToken,
        int timeoutMs = 1_500, int pollMs = 20)
    {
        var deadline = Environment.TickCount64 + timeoutMs;
        nint candidate = 0;
        string? candidateId = null;
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

                // Lost/extra tabs mean another operation is competing with this click.
                if (!before.Handles.All(current.Tabs.Handles.Contains) ||
                    current.Tabs.Handles.Length > before.Handles.Length + 1)
                    return NavigationActivationResult.Ambiguous;

                if (TryGetOnlyAddition(before.Handles, current.Tabs.Handles, out var newHandle) &&
                    TryGetOnlyAddition(before.AutomationIds, current.Tabs.AutomationIds, out var newId))
                {
                    if (candidate != 0 && (candidate != newHandle || candidateId != newId))
                        return NavigationActivationResult.Ambiguous;
                    if (current.ActiveTab == newHandle)
                        return selected ? NavigationActivationResult.Activated : NavigationActivationResult.AlreadyActive;
                    if (current.ActiveTab != sourceTab)
                    {
                        ExplorerDebugLog.Write($"Navigation active tab changed source={sourceTab} target={newHandle} active={current.ActiveTab}");
                        return NavigationActivationResult.Cancelled;
                    }

                    if (candidate == 0)
                    {
                        candidate = newHandle;
                        candidateId = newId;
                        confirmedAt = Environment.TickCount64;
                    }
                    // Require a second coherent observation; late competing creations must not be guessed.
                    else if (!selected && Environment.TickCount64 - confirmedAt >= 40)
                    {
                        if (cancellationToken.IsCancellationRequested || !isCurrent())
                            return NavigationActivationResult.Cancelled;
                        if (!select(newId))
                            return NavigationActivationResult.SelectionRejected;
                        selected = true;
                    }
                }
                else
                {
                    if (selected || candidate != 0 || current.ActiveTab != sourceTab)
                        return NavigationActivationResult.Ambiguous;
                    // Native HWNDs can precede UIA publication. Wait, but never select a partial match.
                    if (current.Tabs.AutomationIds.Length > before.AutomationIds.Length + 1)
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
