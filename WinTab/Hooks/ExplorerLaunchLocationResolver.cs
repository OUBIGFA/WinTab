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
        CancellationToken cancellationToken = default)
    {
        var location = Normalize(readLocation());
        if (!IsStartup(location, isStartupLocation))
            return await WaitForStableLocationAsync(location, readLocation, cancellationToken);

        var fallbackLocation = location;
        var deadline = Environment.TickCount64 + _options.DefaultLocationWaitMs;
        while (!cancellationToken.IsCancellationRequested && Environment.TickCount64 < deadline)
        {
            await Task.Delay(_options.PollIntervalMs, cancellationToken);

            location = Normalize(readLocation());
            if (string.IsNullOrWhiteSpace(location))
                continue;

            fallbackLocation = location;
            if (!IsStartup(location, isStartupLocation))
                return await WaitForStableLocationAsync(location, readLocation, cancellationToken);
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

    public sealed record Options(
        int DefaultLocationWaitMs = 500,
        int StableLocationWaitMs = 160,
        int PollIntervalMs = 25);
}
