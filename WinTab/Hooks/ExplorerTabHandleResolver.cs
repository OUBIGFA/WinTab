using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using WinTab.Helpers;

namespace WinTab.Hooks;

internal static class ExplorerTabHandleResolver
{
    public static async Task<nint> WaitAsync(Func<Task<nint>> getHandle, int timeoutMs, int pollSleepMs)
    {
        async Task<nint> ReadHandleAsync()
        {
            try
            {
                return await getHandle();
            }
            catch (COMException exception)
            {
                ExplorerDebugLog.Write($"Tab handle query not ready error={exception.GetType().Name}:{exception.Message}");
                return 0;
            }
        }

        try
        {
            return await Helper.DoUntilNotDefaultAsync(ReadHandleAsync, timeoutMs, pollSleepMs);
        }
        catch (TimeoutException ex)
        {
            ExplorerDebugLog.Write($"Tab handle resolution failed error={ex.GetType().Name}:{ex.Message}");
            return 0;
        }
    }
}
