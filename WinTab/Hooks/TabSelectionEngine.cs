using System;
using System.Threading;
using System.Threading.Tasks;

namespace WinTab.Hooks;

internal static class TabSelectionEngine
{
    public static async Task<bool> CycleToTabAsync(
        nint targetTab,
        Func<nint[]> getAllTabs,
        Func<nint> getActiveTab,
        Action<int> sendSelectByIndex,
        int totalTimeoutMs = 500,
        int pollSleepMs = 5,
        int perStepTimeoutMs = 0,
        CancellationToken cancellationToken = default)
    {
        if (targetTab == 0) return false;

        var tabs = getAllTabs();
        if (tabs.Length == 0) return false;

        if (getActiveTab() == targetTab) return true;

        var perStep = perStepTimeoutMs > 0
            ? perStepTimeoutMs
            : Math.Max(20, totalTimeoutMs / Math.Max(1, tabs.Length));

        var startedAt = Environment.TickCount64;
        var globalDeadline = startedAt + totalTimeoutMs;

        for (var i = 0; i < tabs.Length; i++)
        {
            if (cancellationToken.IsCancellationRequested)
                return false;

            var activeTab = getActiveTab();
            if (activeTab == targetTab)
                return true;

            if (Environment.TickCount64 >= globalDeadline)
                return getActiveTab() == targetTab;

            sendSelectByIndex(i);

            var previousTab = activeTab;
            var stepDeadline = Math.Min(globalDeadline, Environment.TickCount64 + perStep);
            while (Environment.TickCount64 < stepDeadline)
            {
                if (cancellationToken.IsCancellationRequested)
                    return false;

                var observed = getActiveTab();
                if (observed == targetTab)
                    return true;

                if (observed != previousTab && observed != 0)
                    break;

                await Task.Delay(pollSleepMs, cancellationToken).ConfigureAwait(false);
            }
        }

        return getActiveTab() == targetTab;
    }
}
