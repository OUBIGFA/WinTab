using System;
using System.Threading;
using System.Threading.Tasks;

namespace WinTab.Hooks;

public sealed class ExplorerLaunchLocationResolver
{
    private readonly Options _options;

    public ExplorerLaunchLocationResolver(Options? options = null)
    {
        _options = options ?? new Options();
    }

    public async Task<string> ResolveAsync(
        Func<string?> readLocation,
        Predicate<string> isStartupLocation,
        CancellationToken cancellationToken = default,
        Func<string, Task>? onStartupLocationRetained = null,
        Func<bool>? isBusy = null)
    {
        var location = Normalize(readLocation());
        if (!IsStartup(location, isStartupLocation))
            return await WaitForStableLocationAsync(location, readLocation, cancellationToken);

        var fallbackLocation = location;
        var releaseDeadline = Environment.TickCount64 + _options.DefaultLocationWaitMs;
        var deadline = Environment.TickCount64 + Math.Max(_options.DefaultLocationWaitMs, _options.MaximumStartupLocationWaitMs);
        var busyDeadline = deadline + _options.BusyStartupLocationWaitMs;
        var notifiedStartupLocationRetained = false;
        while (!cancellationToken.IsCancellationRequested)
        {
            var now = Environment.TickCount64;
            if (now >= deadline &&
                (isBusy == null || !SafeIsBusy(isBusy) || now >= busyDeadline))
            {
                break;
            }

            await Task.Delay(_options.PollIntervalMs, cancellationToken);

            location = Normalize(readLocation());
            if (string.IsNullOrWhiteSpace(location))
            {
                if (!notifiedStartupLocationRetained &&
                    onStartupLocationRetained != null &&
                    Environment.TickCount64 >= releaseDeadline)
                {
                    notifiedStartupLocationRetained = true;
                    await onStartupLocationRetained(fallbackLocation);
                }

                continue;
            }

            fallbackLocation = location;
            if (!IsStartup(location, isStartupLocation))
                return await WaitForStableLocationAsync(location, readLocation, cancellationToken);

            if (!notifiedStartupLocationRetained &&
                onStartupLocationRetained != null &&
                Environment.TickCount64 >= releaseDeadline)
            {
                notifiedStartupLocationRetained = true;
                await onStartupLocationRetained(location);
            }
        }

        return fallbackLocation;
    }

    private async Task<string> WaitForStableLocationAsync(
        string firstLocation,
        Func<string?> readLocation,
        CancellationToken cancellationToken)
    {
        var stableLocation = firstLocation;
        var deadline = Environment.TickCount64 + _options.StableLocationWaitMs;
        while (!cancellationToken.IsCancellationRequested && Environment.TickCount64 < deadline)
        {
            await Task.Delay(_options.PollIntervalMs, cancellationToken);

            var currentLocation = Normalize(readLocation());
            if (string.IsNullOrWhiteSpace(currentLocation) ||
                StringComparer.OrdinalIgnoreCase.Equals(currentLocation, stableLocation))
                continue;

            stableLocation = currentLocation;
            deadline = Environment.TickCount64 + _options.StableLocationWaitMs;
        }

        return stableLocation;
    }

    private static bool IsStartup(string location, Predicate<string> isStartupLocation)
    {
        return string.IsNullOrWhiteSpace(location) || isStartupLocation(location);
    }

    private static string Normalize(string? location)
    {
        return string.IsNullOrWhiteSpace(location) ? string.Empty : location.Trim();
    }

    private static bool SafeIsBusy(Func<bool> isBusy)
    {
        try
        {
            return isBusy();
        }
        catch
        {
            return false;
        }
    }

    public sealed record Options(
        int DefaultLocationWaitMs = 60,
        int StableLocationWaitMs = 160,
        int PollIntervalMs = 25,
        int MaximumStartupLocationWaitMs = 500,
        int BusyStartupLocationWaitMs = 1_500);
}
