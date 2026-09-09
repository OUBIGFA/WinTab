using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using WinTab.Hooks;

internal static class ExplorerTabReuseTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("cached tab matches without reading other COM locations", CachedMatchAvoidsUnneededReads);
        yield return ("cached tab reuse survives repeated unavailable Explorer reads", CachedMatchSurvivesUnavailableReads);
        yield return ("stale cached location falls back to the live location", StaleCacheFallsBackToLiveLocation);
        yield return ("different locations are not misclassified as a reuse match", DifferentLocationsDoNotMatch);
        yield return ("a disconnected tab read is not reported as a missing tab", DisconnectedReadIsNotAMiss);
        yield return ("a disconnected location comparison is not reported as a missing tab", DisconnectedComparisonIsNotAMiss);
        yield return ("a single tab publishes its active handle immediately", SingleTabPublishesImmediately);
        yield return ("multiple tabs do not use the active handle as a shortcut", MultipleTabsRequireExactResolution);
    }

    private static Task CachedMatchAvoidsUnneededReads()
    {
        var reads = new[] { 0, 0, 0 };
        var candidates = new[]
        {
            new ExplorerTabReuseCandidate(101, @"C:\Work", () =>
            {
                reads[0]++;
                return @"C:\Work";
            }),
            new ExplorerTabReuseCandidate(202, @"C:\Other", () =>
            {
                reads[1]++;
                return @"C:\Other";
            }),
            new ExplorerTabReuseCandidate(303, @"C:\Archive", () =>
            {
                reads[2]++;
                return @"C:\Archive";
            })
        };

        var found = ExplorerTabReuseMatcher.TryFind(
            @"c:\work",
            candidates,
            AreSameLocation,
            out var handle);

        Check.That(found, "A matching cached location must be reusable.");
        Check.Equal((nint)101, handle, "The cached matching tab must be selected.");
        Check.Equal(0, reads[0], "A known matching tab must not wait for another COM location read.");
        Check.Equal(0, reads[1], "Nonmatching tabs must not incur a COM location read on a cache hit.");
        Check.Equal(0, reads[2], "Nonmatching tabs must not incur a COM location read on a cache hit.");
        return Task.CompletedTask;
    }

    private static Task CachedMatchSurvivesUnavailableReads()
    {
        foreach (var location in new[] { @"C:\Work", @"\\server\share\Archive" })
        {
            var candidates = new[]
            {
                new ExplorerTabReuseCandidate(606, location,
                    () => throw new System.Runtime.InteropServices.COMException("Explorer is temporarily unavailable."))
            };

            for (var attempt = 0; attempt < 1_000; attempt++)
            {
                var found = ExplorerTabReuseMatcher.TryFind(location, candidates, AreSameLocation, out var handle);
                Check.That(found, "A transient Explorer read failure must not disable an already known tab.");
                Check.Equal((nint)606, handle, "Repeated reuse must keep selecting the original tab.");
            }
        }

        return Task.CompletedTask;
    }

    private static Task StaleCacheFallsBackToLiveLocation()
    {
        var cachedLocation = @"C:\Old";
        var candidates = new[]
        {
            new ExplorerTabReuseCandidate(
                404,
                cachedLocation,
                () => @"C:\Current",
                location => cachedLocation = location)
        };

        var found = ExplorerTabReuseMatcher.TryFind(
            @"C:\Current",
            candidates,
            AreSameLocation,
            out var handle);

        Check.That(found, "A stale cache must not prevent finding the tab's current location.");
        Check.Equal((nint)404, handle, "The live matching tab must be selected.");
        Check.Equal(@"C:\Current", cachedLocation, "A successful live fallback must refresh the cache.");
        return Task.CompletedTask;
    }

    private static Task DifferentLocationsDoNotMatch()
    {
        var candidates = new[]
        {
            new ExplorerTabReuseCandidate(505, @"C:\Other", () => @"C:\Other")
        };

        var found = ExplorerTabReuseMatcher.TryFind(
            @"C:\Target",
            candidates,
            AreSameLocation,
            out var handle);

        Check.That(!found, "Different folders must not be reported as the same tab.");
        Check.Equal((nint)0, handle, "A nonmatching search must not return a tab handle.");
        return Task.CompletedTask;
    }

    private static Task DisconnectedReadIsNotAMiss()
    {
        var candidate = new ExplorerTabReuseCandidate(404, null,
            () => throw new COMException("Disconnected", unchecked((int)0x80010108)));
        Check.Throws<COMException>(() => ExplorerTabReuseMatcher.TryFind(@"C:\Work", [candidate], AreSameLocation, out _),
            "A broken connection must reach recovery instead of authorizing a duplicate tab.");
        return Task.CompletedTask;
    }

    private static Task DisconnectedComparisonIsNotAMiss()
    {
        var candidate = new ExplorerTabReuseCandidate(505, @"C:\Work", () => @"C:\Work");
        Check.Throws<InvalidComObjectException>(() => ExplorerTabReuseMatcher.TryFind(@"C:\Work", [candidate],
            (_, _) => throw new InvalidComObjectException("Released connection"), out _),
            "An unavailable comparer cannot confirm that there are no reusable tabs.");
        return Task.CompletedTask;
    }

    private static Task SingleTabPublishesImmediately()
    {
        var activeReads = 0;
        var published = ExplorerTabHandlePublisher.TryGetSingleTabHandle(
            [707],
            () =>
            {
                activeReads++;
                return 707;
            },
            out var handle);

        Check.That(published, "A single tab with a valid active handle can be published immediately.");
        Check.Equal((nint)707, handle, "The single tab handle must be returned.");
        Check.Equal(1, activeReads, "The active handle should be read once.");
        return Task.CompletedTask;
    }

    private static Task MultipleTabsRequireExactResolution()
    {
        var activeReads = 0;
        var published = ExplorerTabHandlePublisher.TryGetSingleTabHandle(
            [808, 909],
            () =>
            {
                activeReads++;
                return 808;
            },
            out var handle);

        Check.That(!published, "Multiple tabs must not be assigned through the single-tab shortcut.");
        Check.Equal((nint)0, handle, "The shortcut must not return an ambiguous handle.");
        Check.Equal(0, activeReads, "The shortcut should reject multiple tabs before reading the active handle.");
        return Task.CompletedTask;
    }

    private static bool AreSameLocation(string left, string right)
    {
        left = left.TrimEnd('\\');
        right = right.TrimEnd('\\');
        return StringComparer.OrdinalIgnoreCase.Equals(left, right);
    }
}
