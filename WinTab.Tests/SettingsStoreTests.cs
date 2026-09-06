using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualBasic.FileIO;
using WinTab.Managers;

internal static class SettingsStoreTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("settings reads and changes do not wait for a blocked disk write", BlockedWriteDoesNotBlockSettings);
        yield return ("settings writes coalesce changes without concurrent file access", WritesLatestSnapshot);
        yield return ("settings save failure is reported without losing in-memory state", ReportsWriteFailure);
        yield return ("settings retain a valid backup and recover from a damaged primary file", RecoversBackup);
        yield return ("failed settings replacement preserves the previous file", FailedReplacementPreservesFile);
    }

    private static string NewPath() => Path.Combine(Path.GetTempPath(), "WinTab.Tests", Guid.NewGuid().ToString("N"), "settings.json");

    private static async Task BlockedWriteDoesNotBlockSettings()
    {
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        using var store = new SettingsStore(NewPath(), settings =>
        {
            started.Set();
            release.Wait();
        });
        var update = Task.Run(() => store.Update(settings => settings with { DoubleClickCloseTab = false }));
        Task<bool>? read = null;
        var responsive = false;
        try
        {
            Check.That(started.Wait(2_000), "The file writer should start.");
            read = Task.Run(() =>
            {
                store.Update(settings => settings with { ReuseTabs = false });
                return store.Snapshot.DoubleClickCloseTab;
            });
            responsive = await Task.WhenAny(read, Task.Delay(300)) == read;
        }
        finally
        {
            release.Set();
            await update;
            if (read != null)
                await read;
            await store.FlushAsync();
        }
        Check.That(responsive, "Mouse settings reads and subsequent edits must not wait for disk I/O.");
        Check.That(!store.Snapshot.DoubleClickCloseTab && !store.Snapshot.ReuseTabs, "The latest state must be visible immediately.");
    }

    private static async Task WritesLatestSnapshot()
    {
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var writes = new List<AppSettings>();
        using var store = new SettingsStore(NewPath(), settings =>
        {
            lock (writes)
                writes.Add(settings);
            started.Set();
            release.Wait();
        });
        store.Update(settings => settings with { Theme = "Dark" });
        Check.That(started.Wait(2_000), "The first write should start.");
        Task<bool> flushed;
        try
        {
            for (var index = 0; index < 1_000; index++)
                store.Update(settings => settings with { Language = index.ToString() }, deferred: true);
            flushed = store.FlushAsync();
            Check.That(!flushed.IsCompleted, "An exit flush must wait for the active disk write and the latest snapshot.");
        }
        finally
        {
            release.Set();
        }
        Check.That(await flushed, "The final write should succeed.");
        Check.That(writes.Count <= 2, "Only the current write and the newest pending snapshot should reach disk.");
        Check.Equal("999", writes[^1].Language);
        Check.Equal("Dark", writes[^1].Theme);
    }

    private static async Task ReportsWriteFailure()
    {
        using var store = new SettingsStore(NewPath(), settings => throw new IOException("Disk unavailable"));
        store.Update(settings => settings with { AutoUpdate = false });
        Check.That(!await store.FlushAsync(), "A failed save must not report success.");
        Check.That(store.LastError is IOException, "The save error must be available to the interface.");
        Check.That(!store.Snapshot.AutoUpdate, "Save failure must not roll back the active setting.");
    }

    private static async Task RecoversBackup()
    {
        var path = NewPath();
        try
        {
            using (var store = new SettingsStore(path))
            {
                store.Update(settings => settings with { DoubleClickCloseTab = false, Language = "简体中文-日本語" });
                Check.That(await store.FlushAsync(), "The initial save should succeed.");
                store.Update(settings => settings with { DoubleClickCloseTab = true });
                Check.That(await store.FlushAsync(), "The replacement should succeed.");
            }
            File.WriteAllText(path, "{broken", Encoding.UTF8);
            using var recovered = new SettingsStore(path);
            Check.That(!recovered.Snapshot.DoubleClickCloseTab, "Corruption should recover the last valid backup.");
            Check.Equal("简体中文-日本語", recovered.Snapshot.Language);
            Check.That(recovered.LastError != null, "Recovery must disclose that the primary file was damaged.");
            Check.That(await recovered.FlushAsync(), "Recovered settings should be writable.");
            var backup = JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(path + ".bak", Encoding.UTF8))!;
            Check.That(!backup.DoubleClickCloseTab, "Repair must not overwrite the valid backup with corrupt data.");
        }
        finally
        {
            RecycleDirectory(path);
        }
    }

    private static async Task FailedReplacementPreservesFile()
    {
        var path = NewPath();
        try
        {
            using var store = new SettingsStore(path);
            Check.That(await store.FlushAsync(), "The initial file should be created.");
            using (var lockedFile = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                store.Update(settings => settings with { AutoUpdate = false });
                Check.That(!await store.FlushAsync(), "Replacing a locked settings file must report failure.");
            }
            var persisted = JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(path, Encoding.UTF8))!;
            Check.That(persisted.AutoUpdate, "A failed replacement must leave the previous settings intact.");
        }
        finally
        {
            RecycleDirectory(path);
        }
    }

    private static void RecycleDirectory(string path)
    {
        var directory = Path.GetDirectoryName(path)!;
        if (Directory.Exists(directory))
            FileSystem.DeleteDirectory(directory, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
    }
}
