using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WinTab.Hooks;

internal static class ExplorerLaunchLocationResolverTests
{
    private const string ThisPc = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
    private const string Downloads = "file:///C:/Users/Public/Downloads";
    private const string TargetFolder = "file:///E:/WinTabStress/Target";

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("resolver waits for the real folder when Explorer first reports This PC", WaitsForRealFolderAfterTransientDefault);
        yield return ("resolver waits longer for delayed external Shell folder launches", WaitsForDelayedExternalShellFolderAfterDefault);
        yield return ("resolver returns the default folder only after the startup location stays default", ReturnsDefaultAfterTimeout);
        yield return ("resolver keeps waiting while a startup-location window is busy", BusyStartupLocationWaitsForRealFolder);
        yield return ("resolver waits for a non-default location to stabilize", WaitsForStableNonDefaultLocation);
        yield return ("resolver returns a stable non-default location immediately", ReturnsStableNonDefaultLocationQuickly);
    }

    private static async Task WaitsForRealFolderAfterTransientDefault()
    {
        var samples = new Queue<string>([ThisPc, ThisPc, TargetFolder, TargetFolder]);

        var resolved = await CreateFastResolver().ResolveAsync(
            () => samples.Count > 0 ? samples.Dequeue() : TargetFolder,
            IsDefaultLocation);

        Check.EqualIgnoreCase(TargetFolder, resolved);
    }

    private static async Task ReturnsDefaultAfterTimeout()
    {
        var resolved = await CreateFastResolver().ResolveAsync(() => ThisPc, IsDefaultLocation);

        Check.EqualIgnoreCase(ThisPc, resolved);
    }

    private static async Task WaitsForDelayedExternalShellFolderAfterDefault()
    {
        var start = Environment.TickCount64;
        var releasedStartupLocation = false;

        var resolved = await CreateFastResolver().ResolveAsync(
            () => Environment.TickCount64 - start < 95 ? ThisPc : TargetFolder,
            IsDefaultLocation,
            onStartupLocationRetained: _ =>
            {
                releasedStartupLocation = true;
                return Task.CompletedTask;
            });

        Check.That(releasedStartupLocation,
            "A merge source that still reports This PC after the short wait must be released instead of kept hidden.");
        Check.EqualIgnoreCase(TargetFolder, resolved);
    }

    private static async Task WaitsForStableNonDefaultLocation()
    {
        var samples = new Queue<string>([ThisPc, Downloads, TargetFolder, TargetFolder]);

        var resolved = await CreateFastResolver().ResolveAsync(
            () => samples.Count > 0 ? samples.Dequeue() : TargetFolder,
            IsDefaultLocation);

        Check.EqualIgnoreCase(TargetFolder, resolved);
    }

    private static async Task BusyStartupLocationWaitsForRealFolder()
    {
        var start = Environment.TickCount64;

        var resolved = await CreateFastResolver().ResolveAsync(
            () => Environment.TickCount64 - start < 95 ? ThisPc : TargetFolder,
            IsDefaultLocation,
            isBusy: () => Environment.TickCount64 - start < 95);

        Check.EqualIgnoreCase(TargetFolder, resolved);
    }

    private static async Task ReturnsStableNonDefaultLocationQuickly()
    {
        var start = Environment.TickCount64;
        var resolved = await CreateFastResolver().ResolveAsync(() => TargetFolder, IsDefaultLocation);
        var elapsed = Environment.TickCount64 - start;

        Check.EqualIgnoreCase(TargetFolder, resolved);
        Check.That(elapsed < 120, $"A stable real folder must resolve within the short stabilization window; took {elapsed} ms.");
    }

    private static ExplorerLaunchLocationResolver CreateFastResolver()
    {
        return new ExplorerLaunchLocationResolver(new ExplorerLaunchLocationResolver.Options(
            DefaultLocationWaitMs: 80,
            StableLocationWaitMs: 15,
            PollIntervalMs: 1,
            MaximumStartupLocationWaitMs: 160));
    }

    private static bool IsDefaultLocation(string location) => StringComparer.OrdinalIgnoreCase.Equals(location, ThisPc);
}
