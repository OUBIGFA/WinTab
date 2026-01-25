using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace WinTab.Helpers;

internal static class ComGuard
{
    public static T? Try<T>(Func<T?> action, string context, Action<string>? onError = null,
        int retryAttempts = Constants.ComRetryAttempts, int retryDelayMs = Constants.ComRetryDelayMs)
    {
        for (var attempt = 0; attempt <= retryAttempts; attempt++)
        {
            try
            {
                return action();
            }
            catch (COMException ex)
            {
                onError?.Invoke($"COMException in {context} (attempt {attempt + 1}): 0x{ex.HResult:X8}");
                if (attempt >= retryAttempts) return default;
                Thread.Sleep(retryDelayMs);
            }
            catch (Exception ex)
            {
                onError?.Invoke($"Exception in {context} (attempt {attempt + 1}): {ex.GetType().Name}");
                if (attempt >= retryAttempts) return default;
                Thread.Sleep(retryDelayMs);
            }
        }

        return default;
    }

    public static bool Try(Action action, string context, Action<string>? onError = null,
        int retryAttempts = Constants.ComRetryAttempts, int retryDelayMs = Constants.ComRetryDelayMs)
    {
        for (var attempt = 0; attempt <= retryAttempts; attempt++)
        {
            try
            {
                action();
                return true;
            }
            catch (COMException ex)
            {
                onError?.Invoke($"COMException in {context} (attempt {attempt + 1}): 0x{ex.HResult:X8}");
                if (attempt >= retryAttempts) return false;
                Thread.Sleep(retryDelayMs);
            }
            catch (Exception ex)
            {
                onError?.Invoke($"Exception in {context} (attempt {attempt + 1}): {ex.GetType().Name}");
                if (attempt >= retryAttempts) return false;
                Thread.Sleep(retryDelayMs);
            }
        }

        return false;
    }
}


