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
        yield return ("settings keep tab restoration off and strict for new configurations", RestoreTabsDefaults);
        yield return ("settings keep tab restoration off and strict for existing configurations", RestoreTabsLegacyDefaults);
        yield return ("settings persist disabled strict tab restoration", () => RestoreTabsPersists(false, false));
        yield return ("settings persist enabled strict tab restoration", () => RestoreTabsPersists(true, false));
        yield return ("settings persist enabled any-folder tab restoration", () => RestoreTabsPersists(true, true));
        yield return ("settings remember any-folder mode while tab restoration is disabled", () => RestoreTabsPersists(false, true));
        yield return ("settings keep single-tab restoration off when upgrading an enabled restore configuration", RestoreSingleTabUpgradeDefault);
        yield return ("settings persist enabled single-tab restoration in strict mode", () => RestoreTabsPersists(true, false, true));
        yield return ("settings persist enabled single-tab restoration in any-folder mode", () => RestoreTabsPersists(true, true, true));
        yield return ("settings remember single-tab restoration while group restoration is disabled", () => RestoreTabsPersists(false, false, true));
        yield return ("settings remember single-tab and any-folder preferences while group restoration is disabled", () => RestoreTabsPersists(false, true, true));
        yield return ("settings enable navigation middle-click for existing configurations independently", MiddleClickDefaults);
        yield return ("settings persist disabling navigation middle-click independently", MiddleClickPersists);
        yield return ("settings reads and changes do not wait for a blocked disk write", BlockedWriteDoesNotBlockSettings);
        yield return ("settings writes coalesce changes without concurrent file access", WritesLatestSnapshot);
        yield return ("settings save failure is reported without losing in-memory state", ReportsWriteFailure);
        yield return ("settings retain a valid backup and recover from a damaged primary file", RecoversBackup);
        yield return ("settings recover a backup after a negative width", () => RecoversInvalidValues("""{"FormSize":{"Width":-1,"Height":720}}"""));
        yield return ("settings recover a backup after a negative height", () => RecoversInvalidValues("""{"FormSize":{"Width":1020,"Height":-1}}"""));
        yield return ("settings recover a backup after a zero dimension", () => RecoversInvalidValues("""{"FormSize":{"Width":1020,"Height":0}}"""));
        yield return ("settings recover a backup after a nonfinite dimension", () => RecoversInvalidValues("""{"FormSize":{"Width":1e999,"Height":720}}"""));
        yield return ("invalid settings and backup use valid defaults and retain the error", InvalidBackupUsesDefaults);
        yield return ("settings without a saved window size leave the size to fit the content", () => LoadsFormSize("""{"Theme":"Dark"}""", null));
        yield return ("settings treat the legacy fixed window size as never chosen", () => LoadsFormSize("""{"FormSize":{"Width":960,"Height":760}}""", null));
        yield return ("settings keep a window size the user chose", () => LoadsFormSize("""{"FormSize":{"Width":960,"Height":900}}""", new System.Windows.Size(960, 900)));
        yield return ("failed settings replacement preserves the previous file", FailedReplacementPreservesFile);
        yield return ("an unreadable primary with a backup cannot be overwritten by fallback settings", () => UnreadablePrimaryIsPreserved(true));
        yield return ("an unreadable primary without a backup cannot be overwritten by defaults", () => UnreadablePrimaryIsPreserved(false));
        yield return ("an unreadable backup is preserved when the primary is damaged", () => UnreadableBackupIsPreserved(true));
        yield return ("an unreadable backup is preserved when the primary is missing", () => UnreadableBackupIsPreserved(false));
    }

    private static string NewPath() => Path.Combine(Path.GetTempPath(), "WinTab.Tests", Guid.NewGuid().ToString("N"), "settings.json");

    private static Task RestoreTabsDefaults()
    {
        using var store = new SettingsStore(NewPath());
        Check.That(!store.Snapshot.RestoreTabs, "Tab restoration must be opt-in on a new installation.");
        Check.That(!store.Snapshot.RestoreOnAnyFolder, "Normal-launch-only mode must be the default.");
        Check.That(!store.Snapshot.RestoreSingleTab, "Single-tab restoration must be off by default.");
        return Task.CompletedTask;
    }

    private static Task RestoreTabsLegacyDefaults()
    {
        var path = NewPath();
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, """{"WindowHook":false,"ReuseTabs":false,"Theme":"Dark"}""");
            using var store = new SettingsStore(path);
            Check.That(store.LastError is null, "Configurations without restoration settings must remain valid.");
            Check.That(!store.Snapshot.RestoreTabs && !store.Snapshot.RestoreOnAnyFolder && !store.Snapshot.RestoreSingleTab,
                "An upgrade must not enable restoration, single-tab restoration or the broader trigger mode.");
            Check.That(!store.Snapshot.WindowHook && !store.Snapshot.ReuseTabs,
                "Loading restoration defaults must not enable merging or reuse.");
            Check.Equal("Dark", store.Snapshot.Theme, "Existing preferences must be retained.");
        }
        finally { RecycleDirectory(path); }
        return Task.CompletedTask;
    }

    private static Task RestoreSingleTabUpgradeDefault()
    {
        var path = NewPath();
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, """{"RestoreTabs":true,"RestoreOnAnyFolder":true}""");
            using var store = new SettingsStore(path);
            Check.That(store.LastError is null, "Existing restoration settings must remain valid.");
            Check.That(!store.Snapshot.RestoreSingleTab, "Upgrading must not opt into single-tab restoration.");
            Check.That(store.Snapshot.RestoreTabs && store.Snapshot.RestoreOnAnyFolder,
                "The selected group restoration mode must remain unchanged.");
        }
        finally { RecycleDirectory(path); }
        return Task.CompletedTask;
    }

    private static async Task RestoreTabsPersists(bool restoreTabs, bool restoreOnAnyFolder, bool restoreSingleTab = false)
    {
        var path = NewPath();
        try
        {
            using (var store = new SettingsStore(path))
            {
                store.Update(settings => settings with
                {
                    RestoreTabs = !restoreTabs, RestoreOnAnyFolder = !restoreOnAnyFolder, RestoreSingleTab = !restoreSingleTab,
                    WindowHook = false, ReuseTabs = false
                });
                Check.That(await store.FlushAsync(), "The initial restoration preferences must save successfully.");
                store.Update(settings => settings with
                {
                    RestoreTabs = restoreTabs, RestoreOnAnyFolder = restoreOnAnyFolder, RestoreSingleTab = restoreSingleTab
                });
                Check.That(await store.FlushAsync(), "Changed restoration preferences must save successfully.");
            }
            using var reloaded = new SettingsStore(path);
            Check.That(reloaded.LastError is null, "Restoration settings must reload without using a backup.");
            Check.Equal(restoreTabs, reloaded.Snapshot.RestoreTabs, "The restoration toggle must survive restarting.");
            Check.Equal(restoreOnAnyFolder, reloaded.Snapshot.RestoreOnAnyFolder,
                "The selected mode must persist independently of the restoration toggle.");
            Check.Equal(restoreSingleTab, reloaded.Snapshot.RestoreSingleTab,
                "Single-tab restoration must persist independently of the group toggle and trigger mode.");
            Check.That(!reloaded.Snapshot.WindowHook && !reloaded.Snapshot.ReuseTabs,
                "Saving restoration preferences must not enable merging or reuse.");
        }
        finally { RecycleDirectory(path); }
    }

    private static Task MiddleClickDefaults()
    {
        var path = NewPath();
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, """{"WindowHook":false,"ReuseTabs":false,"DoubleClickCloseTab":false}""");
            using var store = new SettingsStore(path);
            Check.That(store.Snapshot.MiddleClickForegroundTab, "An existing configuration should enable the new setting by default.");
            Check.That(!store.Snapshot.WindowHook && !store.Snapshot.ReuseTabs && !store.Snapshot.DoubleClickCloseTab,
                "The new feature must not turn on merging, reuse, or double-click closing.");
        }
        finally { RecycleDirectory(path); }
        return Task.CompletedTask;
    }

    private static async Task MiddleClickPersists()
    {
        var path = NewPath();
        try
        {
            using (var store = new SettingsStore(path))
            {
                store.Update(settings => settings with { MiddleClickForegroundTab = false });
                Check.That(await store.FlushAsync(), "Disabling middle-click activation must save successfully.");
            }
            using var reloaded = new SettingsStore(path);
            Check.That(!reloaded.Snapshot.MiddleClickForegroundTab, "The disabled setting must survive restarting.");
            Check.That(reloaded.Snapshot.WindowHook && reloaded.Snapshot.ReuseTabs && reloaded.Snapshot.DoubleClickCloseTab,
                "Disabling middle-click activation must leave the other features unchanged.");
        }
        finally { RecycleDirectory(path); }
    }

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

    private static async Task RecoversInvalidValues(string json)
    {
        var path = NewPath();
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, json, Encoding.UTF8);
            File.WriteAllText(path + ".bak", """{"Theme":"Dark","FormSize":{"Width":1100,"Height":760}}""", Encoding.UTF8);
            using var recovered = new SettingsStore(path);
            Check.Equal("Dark", recovered.Snapshot.Theme, "Invalid values must recover the backup rather than prevent startup.");
            Check.Equal(1100d, recovered.Snapshot.FormSize?.Width, "The valid backup width must be preserved.");
            Check.Equal(760d, recovered.Snapshot.FormSize?.Height, "The valid backup height must be preserved.");
            Check.That(recovered.LastError is JsonException, "The invalid setting must be reported as damaged data.");
            Check.That(await recovered.FlushAsync(), "Recovered values must be writable.");
            using var reloaded = new SettingsStore(path);
            Check.Equal("Dark", reloaded.Snapshot.Theme, "The repaired primary file must retain the backup values.");
        }
        finally { RecycleDirectory(path); }
    }

    private static Task InvalidBackupUsesDefaults()
    {
        var path = NewPath();
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, """{"FormSize":{"Width":-1,"Height":720}}""", Encoding.UTF8);
            File.WriteAllText(path + ".bak", """{"FormSize":{"Width":1020,"Height":-1}}""", Encoding.UTF8);
            using var recovered = new SettingsStore(path);
            Check.That(recovered.Snapshot.FormSize is null, "An invalid backup must not replace the content-fitted default size.");
            Check.That(recovered.LastError is JsonException, "Using defaults must not hide the damaged-settings error.");
        }
        finally { RecycleDirectory(path); }
        return Task.CompletedTask;
    }

    private static Task LoadsFormSize(string json, System.Windows.Size? expected)
    {
        var path = NewPath();
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, json, Encoding.UTF8);
            using var store = new SettingsStore(path);
            Check.That(store.LastError is null, "A valid settings file must load without an error.");
            Check.That(store.Snapshot.FormSize == expected,
                $"Expected window size {expected?.ToString() ?? "unset"}, got {store.Snapshot.FormSize?.ToString() ?? "unset"}.");
        }
        finally { RecycleDirectory(path); }
        return Task.CompletedTask;
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

    private static async Task UnreadablePrimaryIsPreserved(bool hasBackup)
    {
        var path = NewPath();
        const string primary = """{"Theme":"Dark","Language":"original-primary"}""";
        const string backup = """{"Theme":"Light","Language":"old-backup"}""";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, primary, Encoding.UTF8);
            if (hasBackup) File.WriteAllText(path + ".bak", backup, Encoding.UTF8);
            SettingsStore store;
            using (var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
                store = new SettingsStore(path);
            using (store)
            {
                Check.That(store.LastError is IOException, "A sharing violation must remain visible to the caller.");
                Check.That(!await store.FlushAsync(), "Exit flush must not overwrite an unreadable primary, even after it is unlocked.");
                store.Update(settings => settings with { Language = "in-memory-change" });
                Check.That(!await store.FlushAsync(), "Changes to fallback settings must not replace unknown on-disk values.");
                Check.Equal("in-memory-change", store.Snapshot.Language, "Read protection must not freeze in-memory settings.");
                Check.That(store.LastError is IOException, "Blocked saving must not clear the original read failure.");
            }
            Check.Equal(primary, File.ReadAllText(path, Encoding.UTF8), "The unreadable primary must remain byte-for-byte unchanged.");
            Check.Equal(hasBackup, File.Exists(path + ".bak"), "Saving fallback values must not create or remove a backup.");
            if (hasBackup) Check.Equal(backup, File.ReadAllText(path + ".bak", Encoding.UTF8), "The valid old backup must be preserved.");
            using var reopened = new SettingsStore(path);
            Check.Equal("original-primary", reopened.Snapshot.Language, "A successful restart must load the original primary.");
            reopened.Update(settings => settings with { Language = "saved-after-restart" });
            Check.That(await reopened.FlushAsync(), "Saving may resume after a successful fresh read.");
        }
        finally { RecycleDirectory(path); }
    }

    private static async Task UnreadableBackupIsPreserved(bool damagedPrimary)
    {
        var path = NewPath();
        const string backup = """{"Language":"only-valid-copy"}""";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            if (damagedPrimary) File.WriteAllText(path, "{damaged", Encoding.UTF8);
            File.WriteAllText(path + ".bak", backup, Encoding.UTF8);
            SettingsStore store;
            using (var locked = new FileStream(path + ".bak", FileMode.Open, FileAccess.Read, FileShare.None))
                store = new SettingsStore(path);
            using (store)
            {
                store.Update(settings => settings with { AutoUpdate = false });
                Check.That(!await store.FlushAsync(), "An unreadable recovery copy must block persistence of unverified defaults.");
                Check.That(store.LastError != null, "The failed read must remain observable.");
            }
            Check.Equal(damagedPrimary, File.Exists(path), "An unavailable backup must not cause a new default primary to be written.");
            if (damagedPrimary) Check.Equal("{damaged", File.ReadAllText(path, Encoding.UTF8), "The primary must be untouched while recovery is uncertain.");
            Check.Equal(backup, File.ReadAllText(path + ".bak", Encoding.UTF8), "The only valid copy must not be replaced.");
            using var reopened = new SettingsStore(path);
            Check.Equal("only-valid-copy", reopened.Snapshot.Language, "Recovery must work once the backup can be read.");
            Check.That(await reopened.FlushAsync(), "Confirmed recovery values may repair the primary.");
        }
        finally { RecycleDirectory(path); }
    }

    private static void RecycleDirectory(string path)
    {
        var directory = Path.GetDirectoryName(path)!;
        if (Directory.Exists(directory))
            FileSystem.DeleteDirectory(directory, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
    }
}
