using System;
using System.Threading.Tasks;
using WinTab.Helpers;

namespace WinTab.Hooks;

internal static class ExplorerTabHandleResolver
{
    public static async Task<nint> WaitAsync(Func<Task<nint>> getHandle, int timeoutMs, int pollSleepMs)
    {
        try
        {
            return await Helper.DoUntilNotDefaultAsync(getHandle, timeoutMs, pollSleepMs);
        }
        catch (Exception ex)
        {
            ExplorerDebugLog.Write($"Tab handle resolution failed error={ex.GetType().Name}:{ex.Message}");
            return 0;
        }
    }
}
