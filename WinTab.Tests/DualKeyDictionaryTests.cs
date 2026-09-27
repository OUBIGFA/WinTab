using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WinTab.Models;

internal static class DualKeyDictionaryTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("DualKeyDictionary resolves entries by primary and optional key", ResolvesByBothKeys);
        yield return ("DualKeyDictionary UpdateOptionalKey re-points the optional lookup", UpdateOptionalKeyRepointsLookup);
        yield return ("DualKeyDictionary Remove clears both lookups", RemoveClearsBothLookups);
        yield return ("DualKeyDictionary rejects duplicate keys on Add and tolerates them on TryAdd", RejectsDuplicates);
        yield return ("DualKeyDictionary pair removal requires both key and value", PairRemovalMatchesValue);
        yield return ("DualKeyDictionary CopyTo validates arguments before copying", CopyValidatesArguments);
        yield return ("DualKeyDictionary enumeration yields every entry with its optional key", EnumerationYieldsEntries);
    }

    private static Task ResolvesByBothKeys()
    {
        var dict = new DualKeyDictionary<string, nint?, int>();
        dict.Add("a", 1, 10);
        dict.Add("b", 2);

        Check.That(dict.TryGetValue("a", out int a) && a == 1, "primary lookup must return the stored value");
        Check.That(dict.TryGetValue((nint?)10, out string? primary) && primary == "a", "optional lookup must return the primary key");
        Check.That(!dict.TryGetValue((nint?)99, out string? _), "unknown optional key must not resolve");
        Check.That(dict.TryGetValue("b", out DualKeyEntry<string, nint?, int> entry) && entry.OptionalKey is null && entry.Value == 2,
            "an entry added without an optional key must report a null optional key");
        Check.Equal(2, dict.Count);
        return Task.CompletedTask;
    }

    private static Task UpdateOptionalKeyRepointsLookup()
    {
        var dict = new DualKeyDictionary<string, nint?, int>();
        dict.Add("a", 1);
        dict.UpdateOptionalKey("a", 42);

        Check.That(dict.TryGetValue((nint?)42, out string? primary) && primary == "a", "the new optional key must resolve");

        dict.UpdateOptionalKey("a", 43);
        Check.That(!dict.ContainsOptional(42), "the old optional key must be released");
        Check.That(dict.ContainsOptional(43), "the replacement optional key must be registered");

        dict.Add("b", 2);
        Check.Throws<ArgumentException>(() => dict.UpdateOptionalKey("b", 43),
            "assigning an optional key that belongs to another primary must be rejected");
        Check.Throws<ArgumentException>(() => dict.UpdateOptionalKey("missing", 1),
            "updating a missing primary key must be rejected");
        return Task.CompletedTask;
    }

    private static Task RemoveClearsBothLookups()
    {
        var dict = new DualKeyDictionary<string, nint?, int>();
        dict.Add("a", 1, 10);
        dict.Add("b", 2, 20);

        Check.That(dict.Remove("a"), "removing an existing primary must succeed");
        Check.That(!dict.ContainsPrimary("a") && !dict.ContainsOptional(10), "both lookups must forget the removed entry");

        Check.That(dict.Remove((nint?)20), "removing by optional key must succeed");
        Check.That(!dict.ContainsPrimary("b"), "removing by optional key must also drop the primary entry");
        Check.Equal(0, dict.Count);
        Check.That(!dict.Remove("a"), "removing a missing key must report false");
        return Task.CompletedTask;
    }

    private static Task RejectsDuplicates()
    {
        var dict = new DualKeyDictionary<string, nint?, int>();
        dict.Add("a", 1, 10);

        Check.Throws<ArgumentException>(() => dict.Add("a", 2), "duplicate primary key must throw on Add");
        Check.Throws<ArgumentException>(() => dict.Add("b", 2, 10), "duplicate optional key must throw on Add");
        Check.That(!dict.TryAdd("a", 3), "TryAdd must report false for a duplicate primary key");
        Check.That(!dict.TryAdd("c", 3, 10), "TryAdd must report false for a duplicate optional key");
        Check.That(dict.TryGetValue("a", out int value) && value == 1, "failed inserts must not modify the existing entry");
        return Task.CompletedTask;
    }

    private static Task PairRemovalMatchesValue()
    {
        var dict = new DualKeyDictionary<string, nint?, int>();
        dict.Add("a", 1, 10);
        ICollection<KeyValuePair<string, int>> collection = dict;
        Check.That(!collection.Remove(new("a", 2)), "A stale value must not remove the current entry.");
        Check.That(dict.ContainsPrimary("a") && dict.ContainsOptional(10), "A mismatch must retain both lookups.");
        Check.That(collection.Remove(new("a", 1)), "The matching pair must be removed.");
        Check.That(dict.Count == 0 && !dict.ContainsOptional(10), "Matching removal must clear both lookups.");
        return Task.CompletedTask;
    }

    private static Task CopyValidatesArguments()
    {
        var dict = new DualKeyDictionary<string, nint?, int>();
        var array = new KeyValuePair<string, int>[2];
        Check.Throws<ArgumentNullException>(() => dict.CopyTo(null!, 0), "Null destinations must be rejected.");
        Check.Throws<ArgumentOutOfRangeException>(() => dict.CopyTo(array, -1), "Even an empty dictionary must reject a negative index.");
        dict.Add("a", 1, 10);
        Check.Throws<ArgumentException>(() => dict.CopyTo(array, 2), "Insufficient space must be rejected before writing.");
        dict.CopyTo(array, 1);
        Check.Equal(new KeyValuePair<string, int>("a", 1), array[1], "Valid copying must respect the offset.");
        return Task.CompletedTask;
    }

    private static Task EnumerationYieldsEntries()
    {
        var dict = new DualKeyDictionary<string, nint?, int>();
        dict.Add("a", 1, 10);
        dict.Add("b", 2);

        var entries = ((IEnumerable<DualKeyEntry<string, nint?, int>>)dict).ToList();
        Check.Equal(2, entries.Count);
        var a = entries.Single(e => e.PrimaryKey == "a");
        Check.That(a.Value == 1 && a.OptionalKey == 10, "enumeration must carry the optional key");
        Check.That(entries.Single(e => e.PrimaryKey == "b").OptionalKey is null, "entries without an optional key enumerate as null");
        return Task.CompletedTask;
    }
}
