using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualBasic.FileIO;
using WinTab.Managers;
using WinTab.Models;

internal static class ExplorerSessionStoreTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("session store does not create history merely by loading or flushing", MissingSessionDoesNotWrite);
        yield return ("session store persists and reloads a single-tab window", () => RoundTrip([@"C:\One"], 0));
        yield return ("session store retains order duplicates and the active tab", () => RoundTrip([@"C:\One", @"D:\同名", @"C:\One"], 2));
        yield return ("session store snapshots cannot mutate pending or saved history", DefensiveCopies);
        yield return ("session store serializes concurrent saves and persists the latest window", LatestWriteWins);
        yield return ("session store reports a failed write and retries without losing memory", WriteFailureIsObservable);
        yield return ("session store recovers a corrupt primary without replacing the valid backup", CorruptPrimaryRecoversBackup);
        yield return ("session store retains a valid backup when the primary was deleted", MissingPrimaryRecoversBackup);
        yield return ("session store never overwrites an unreadable primary", UnreadablePrimaryIsPreserved);
        yield return ("session store never overwrites an unreadable backup", UnreadableBackupIsPreserved);
        yield return ("session store refuses oversized files without overwriting them", OversizedFileIsPreserved);
        yield return ("session store rejects a future schema without downgrading it", FutureVersionIsPreserved);
        yield return ("session store preserves future schemas with missing or changed legacy fields", FutureVersionWithChangedFieldsIsPreserved);
        yield return ("session store preserves its old primary after a replacement failure", ReplacementFailurePreservesPrimary);
        foreach (var json in new[]
        {
            "null", "{}", "[]", "{bad", """{"Version":1,"Locations":[],"ActiveTabIndex":0,"OrderVerified":true}""",
            """{"Version":1,"Locations":["C:\\One"],"ActiveTabIndex":-1,"OrderVerified":true}""",
            """{"Version":1,"Locations":["C:\\One"],"ActiveTabIndex":1,"OrderVerified":true}""",
            """{"Version":1,"Locations":[null],"ActiveTabIndex":0,"OrderVerified":true}""",
            """{"Version":1,"Locations":[" "],"ActiveTabIndex":0,"OrderVerified":true}""",
            """{"Version":1,"Locations":["C:\\bad\u0000"],"ActiveTabIndex":0,"OrderVerified":true}"""
        })
            yield return ($"session store rejects malformed history {json}", () => InvalidFileIsReported(json));
        yield return ("session store rejects empty and oversized groups before changing history", InvalidSaveDoesNotReplaceHistory);
    }

    private static ExplorerSession Session(params string[] locations) => new() { Locations = locations, OrderVerified = true };

    private static Task MissingSessionDoesNotWrite() => WithPath(async path =>
    {
        using var store = new ExplorerSessionStore(path);
        Check.That(store.Snapshot == null && store.LastError == null, "A first run must not invent a previous session.");
        Check.That(await store.FlushAsync(), "Nothing to save is a successful no-op.");
        Check.That(!File.Exists(path), "Loading/closing WinTab with no history must not write a history file.");
    });

    private static Task RoundTrip(string[] locations, int active) => WithPath(async path =>
    {
        using (var store = new ExplorerSessionStore(path))
            Check.That(await store.SaveAsync(Session(locations) with { ActiveTabIndex = active }), "The complete group must save.");
        using var loaded = new ExplorerSessionStore(path);
        Check.That(loaded.LastError == null && loaded.Snapshot != null, "The group must load without fallback.");
        Check.That(loaded.Snapshot!.Locations.SequenceEqual(locations), "Paths, order and duplicate tabs must survive reload.");
        Check.Equal(active, loaded.Snapshot.ActiveTabIndex, "The active tab must survive reload, including a duplicate path.");
        Check.That(loaded.Snapshot.OrderVerified, "The order evidence must survive reload.");
    });

    private static Task DefensiveCopies() => WithPath(async path =>
    {
        using var gate = new ManualResetEventSlim();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        ExplorerSession? written = null;
        using var store = new ExplorerSessionStore(path, snapshot =>
        {
            entered.TrySetResult();
            gate.Wait();
            written = snapshot.Copy();
            snapshot.Locations[0] = "writer mutation";
        });
        var source = Session(@"C:\original");
        var save = store.SaveAsync(source);
        try
        {
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(2));
            source.Locations[0] = "caller mutation";
            store.Snapshot!.Locations[0] = "reader mutation";
            Check.Equal(@"C:\original", store.Snapshot!.Locations[0], "Neither input nor output arrays may mutate the stored group.");
        }
        finally { gate.Set(); }
        Check.That(await save, "The pending write must finish.");
        Check.Equal(@"C:\original", written!.Locations[0], "A pending disk snapshot must be isolated too.");
        Check.Equal(@"C:\original", store.Snapshot!.Locations[0], "A writer must not mutate the in-memory snapshot.");
    });

    private static Task LatestWriteWins() => WithPath(async path =>
    {
        using var release = new ManualResetEventSlim();
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var writes = new List<string>();
        var active = 0;
        var peak = 0;
        using var store = new ExplorerSessionStore(path, snapshot =>
        {
            peak = Math.Max(peak, Interlocked.Increment(ref active));
            if (writes.Count == 0) { started.TrySetResult(); release.Wait(); }
            writes.Add(snapshot.Locations[0]);
            Interlocked.Decrement(ref active);
        });
        var first = store.SaveAsync(Session(@"C:\first"));
        Task<bool>? last = null;
        try
        {
            await started.Task.WaitAsync(TimeSpan.FromSeconds(2));
            for (var i = 0; i < 100; i++) last = store.SaveAsync(Session(@"C:\last-" + i));
            Check.Equal(@"C:\last-99", store.Snapshot!.Locations[0], "Slow storage must not delay the latest closed window in memory.");
        }
        finally { release.Set(); }
        Check.That(await first && await last! && await store.FlushAsync(), "All callers must observe the final save result.");
        Check.Equal(1, peak, "The primary and backup must have only one writer.");
        Check.That(writes.SequenceEqual([@"C:\first", @"C:\last-99"]), "Intermediate closes may coalesce, but the newest one must win.");
    });

    private static Task WriteFailureIsObservable() => WithPath(async path =>
    {
        var fail = true;
        var notifications = 0;
        using var store = new ExplorerSessionStore(path, _ => { if (fail) throw new IOException("disk unavailable"); });
        store.ErrorChanged += () => notifications++;
        Check.That(!await store.SaveAsync(Session(@"C:\retained")), "A failed write cannot report success.");
        Check.That(store.LastError is IOException && notifications == 1, "Failure must be observable, not silently swallowed.");
        Check.Equal(@"C:\retained", store.Snapshot!.Locations[0], "Disk failure must not lose the current session in memory.");
        fail = false;
        Check.That(await store.FlushAsync() && store.LastError == null, "A retry must repair the latest revision.");
        Check.Equal(2, notifications, "Successful recovery must clear the error notification.");
    });

    private static Task CorruptPrimaryRecoversBackup() => WithPath(async path =>
    {
        using (var store = new ExplorerSessionStore(path))
        {
            Check.That(await store.SaveAsync(Session(@"C:\backup")), "The first file must save.");
            Check.That(await store.SaveAsync(Session(@"C:\primary")), "The second file must create a backup.");
        }
        var backup = File.ReadAllText(path + ".bak");
        File.WriteAllText(path, "{corrupt");
        using var recovered = new ExplorerSessionStore(path);
        Check.Equal(@"C:\backup", recovered.Snapshot!.Locations[0], "Corruption must recover the last valid complete group.");
        Check.That(recovered.LastError != null, "Recovery must retain evidence of the damaged primary.");
        Check.That(await recovered.SaveAsync(Session(@"C:\new")), "A known corrupt primary may be repaired by a new closed window.");
        Check.Equal(backup, File.ReadAllText(path + ".bak"), "The damaged primary must never overwrite the valid backup.");
    });

    private static Task MissingPrimaryRecoversBackup() => WithPath(async path =>
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path + ".bak", JsonSerializer.Serialize(Session(@"C:\backup")));
        using var store = new ExplorerSessionStore(path);
        Check.Equal(@"C:\backup", store.Snapshot!.Locations[0], "An existing backup must be read even when the primary is missing.");
        Check.That(store.LastError != null, "The fallback must not be silent.");
        Check.That(await store.SaveAsync(Session(@"C:\new")), "A new primary must be writable after confirmed backup recovery.");
        using var loaded = new ExplorerSessionStore(path);
        Check.Equal(@"C:\new", loaded.Snapshot!.Locations[0], "The new primary must load normally.");
    });

    private static Task UnreadablePrimaryIsPreserved() => WithPath(async path =>
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var original = JsonSerializer.Serialize(Session(@"C:\unknown"));
        File.WriteAllText(path, original);
        File.WriteAllText(path + ".bak", JsonSerializer.Serialize(Session(@"C:\backup")));
        using var locked = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var store = new ExplorerSessionStore(path);
        Check.Equal(@"C:\backup", store.Snapshot!.Locations[0], "A known readable backup remains useful in memory.");
        Check.That(!await store.SaveAsync(Session(@"C:\new")), "An unreadable primary must not be overwritten by a fallback.");
        locked.Dispose();
        Check.That(!await store.FlushAsync(), "Unlocking alone must not authorize replacing data this store never read.");
        Check.Equal(original, File.ReadAllText(path), "Unknown primary data must remain unchanged.");
    });

    private static Task UnreadableBackupIsPreserved() => WithPath(async path =>
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, "bad");
        var backup = JsonSerializer.Serialize(Session(@"C:\only-copy"));
        File.WriteAllText(path + ".bak", backup);
        using var locked = new FileStream(path + ".bak", FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var store = new ExplorerSessionStore(path);
        Check.That(!await store.SaveAsync(Session(@"C:\fallback")), "An unreadable only-good-copy must block replacing the pair.");
        locked.Dispose();
        Check.Equal(backup, File.ReadAllText(path + ".bak"), "The unknown backup must remain intact.");
    });

    private static Task OversizedFileIsPreserved() => WithPath(async path =>
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using (var file = File.Create(path)) file.SetLength(ExplorerSessionStore.MaxFileBytes + 1L);
        using var store = new ExplorerSessionStore(path);
        Check.That(store.Snapshot == null && store.LastError != null, "Oversized data must not be read as history.");
        Check.That(!await store.SaveAsync(Session(@"C:\new")), "An oversized unknown file must be preserved.");
        Check.Equal(ExplorerSessionStore.MaxFileBytes + 1L, new FileInfo(path).Length, "The original file must not be truncated.");
    });

    private static Task FutureVersionIsPreserved() => WithPath(async path =>
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var original = JsonSerializer.Serialize(Session(@"C:\future") with { Version = 99 });
        File.WriteAllText(path, original);
        using var store = new ExplorerSessionStore(path);
        Check.That(store.Snapshot == null && store.LastError is JsonException, "A future format must not be misinterpreted.");
        Check.That(!await store.SaveAsync(Session(@"C:\new")), "Saving must not downgrade an unknown format.");
        Check.Equal(original, File.ReadAllText(path), "The future file must remain unchanged.");
    });

    private static Task FutureVersionWithChangedFieldsIsPreserved() => WithPath(async path =>
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var backup = JsonSerializer.Serialize(Session(@"C:\backup"));
        File.WriteAllText(path + ".bak", backup);
        foreach (var original in new[]
        {
            """{"Version":99,"Tabs":["C:\\future"]}""",
            """{"Version":99,"Locations":{"Paths":["C:\\future"]},"ActiveTabIndex":0,"OrderVerified":true}"""
        })
        {
            File.WriteAllText(path, original);
            using var store = new ExplorerSessionStore(path);
            Check.That(store.LastError is ExplorerSession.UnsupportedVersionException,
                "A future schema must be recognized before current fields are deserialized.");
            Check.That(!await store.SaveAsync(Session(@"C:\new")), "A future schema must never be downgraded.");
            Check.Equal(original, File.ReadAllText(path), "The unknown primary must remain byte-for-byte unchanged.");
            Check.Equal(backup, File.ReadAllText(path + ".bak"), "Its backup must remain untouched too.");
        }
    });

    private static Task ReplacementFailurePreservesPrimary() => WithPath(async path =>
    {
        using var store = new ExplorerSessionStore(path);
        Check.That(await store.SaveAsync(Session(@"C:\original")), "The original must save.");
        var original = File.ReadAllText(path);
        using (var locked = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
            Check.That(!await store.SaveAsync(Session(@"C:\new")), "A blocked atomic replace must fail visibly.");
        Check.Equal(original, File.ReadAllText(path), "A failed replacement must leave the primary byte-for-byte intact.");
        Check.That(!Directory.EnumerateFiles(Path.GetDirectoryName(path)!, "*.tmp").Any(), "Only our temporary file may be removed after failure.");
        Check.That(await store.FlushAsync(), "The latest session must be retryable after the file is unlocked.");
    });

    private static Task InvalidFileIsReported(string json) => WithPath(path =>
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, json);
        using var store = new ExplorerSessionStore(path);
        Check.That(store.Snapshot == null && store.LastError is JsonException, "Malformed history must not produce a restore plan.");
        return Task.CompletedTask;
    });

    private static Task InvalidSaveDoesNotReplaceHistory() => WithPath(async path =>
    {
        using var store = new ExplorerSessionStore(path);
        Check.That(await store.SaveAsync(Session(@"C:\valid")), "The valid group must save.");
        Check.Throws<JsonException>(() => store.SaveAsync(Session()), "Empty groups must be rejected.");
        Check.Throws<JsonException>(() => store.SaveAsync(Session(Enumerable.Repeat(@"C:\one", ExplorerSession.MaxTabs + 1).ToArray())),
            "An oversized group must not be silently truncated.");
        Check.Equal(@"C:\valid", store.Snapshot!.Locations[0], "Rejected input must not alter the last valid group.");
    });

    private static async Task WithPath(Func<string, Task> test)
    {
        var directory = Path.Combine(Path.GetTempPath(), "WinTab.Tests", Guid.NewGuid().ToString("N"));
        try { await test(Path.Combine(directory, "session.json")); }
        finally
        {
            if (Directory.Exists(directory))
                FileSystem.DeleteDirectory(directory, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
        }
    }
}
