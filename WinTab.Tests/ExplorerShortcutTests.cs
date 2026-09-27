using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.Managers;

internal static class ExplorerShortcutTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("session shortcuts normalize and validate configurable chords", Parsing);
        yield return ("session tab shortcuts leave browsers alone and suppress autorepeat", Dispatch);
        yield return ("session group shortcuts work globally with default and custom chords", GlobalDispatch);
        yield return ("session group shortcuts own repeats and releases across foreground changes", GlobalRelease);
        yield return ("session shortcuts allow physical chord entry without executing recovery", ShortcutEntry);
        yield return ("session group shortcut delivery does not depend on its originating window", () => QueuedDelivery(false));
        yield return ("session shortcut disposal cancels queued global recovery", () => QueuedDelivery(true));
        yield return ("session shortcuts preserve consumed key-up across focus changes", Release);
        yield return ("session shortcuts ignore injected input and extra modifiers", NonMatching);
        yield return ("session recovery settings default on without enabling automatic restore", Defaults);
        yield return ("session recovery shortcut defaults apply to configurations without shortcut fields", MissingShortcutDefaults);
        yield return ("session recovery default Alt shortcuts survive a storage round-trip", DefaultPersistence);
        yield return ("session recovery shortcut settings survive a storage round-trip", Persistence);
    }
    private static Task Parsing()
    {
        Check.That(ExplorerShortcut.TryParse("shift + control + e", out var parsed), "Modifier order and whitespace are accepted.");
        Check.Equal("Ctrl+Shift+E", parsed.ToString(), "Saved shortcuts have a canonical display.");
        foreach (var valid in new[] { "Ctrl+Alt+7", "Alt+F24", "Ctrl+Shift+F1" })
            Check.That(ExplorerShortcut.TryParse(valid, out _), "Valid chord rejected: " + valid);
        foreach (var invalid in new[] { "E", "Shift+E", "Ctrl+Ctrl+E", "Ctrl+", "Ctrl++E", "Win+E", "Ctrl+F25", "Ctrl+Escape", "Ctrl+65" })
            Check.That(!ExplorerShortcut.TryParse(invalid, out _), "Invalid chord accepted: " + invalid);
        return Task.CompletedTask;
    }
    private static ExplorerShortcutDispatch Controller()
    {
        var settings = new AppSettings();
        ExplorerShortcut.TryParse(settings.RestoreGroupShortcut, out var group);
        ExplorerShortcut.TryParse(settings.ReopenTabShortcut, out var tab);
        return new ExplorerShortcutDispatch { Group = group, Tab = tab };
    }
    private const ShortcutModifiers DefaultModifiers = ShortcutModifiers.Alt;
    private static Task Dispatch()
    {
        var controller = Controller();
        Check.That(!controller.Handle('W', true, DefaultModifiers, false, false, out var browser) && browser == null,
            "Browser Alt+W must pass through.");
        controller.Handle('W', false, DefaultModifiers, false, false, out _);
        Check.That(controller.Handle('W', true, DefaultModifiers, true, false, out var first) && first == SessionAction.ReopenTab,
            "Explorer invokes the command once.");
        Check.That(controller.Handle('W', true, DefaultModifiers, true, false, out var repeat) && repeat == null,
            "Holding the shortcut does not reopen the whole history.");
        Check.That(controller.Handle('W', false, DefaultModifiers, true, false, out _), "Its key-up is consumed.");
        Check.That(controller.Handle('W', true, DefaultModifiers, true, false, out var second) && second == SessionAction.ReopenTab,
            "A later physical press can invoke again.");
        Check.That(!controller.Handle('T', true, ShortcutModifiers.Control | ShortcutModifiers.Shift, false, false, out var browserTab) && browserTab == null,
            "Browser Ctrl+Shift+T still passes through.");
        return Task.CompletedTask;
    }
    private static Task GlobalDispatch()
    {
        foreach (var explorer in new[] { false, true })
        {
            foreach (var text in new[] { new AppSettings().RestoreGroupShortcut, "Ctrl+Shift+E", "Ctrl+Alt+F8", "Alt+7" })
            {
                var controller = Controller();
                ExplorerShortcut.TryParse(text, out var group);
                controller.Group = group;
                Check.That(controller.Handle(group.Key, true, group.Modifiers, explorer, false, out var action) &&
                    action == SessionAction.RestoreGroup, "Group recovery must work in every foreground application: " + text);
                Check.That(controller.Handle(group.Key, false, ShortcutModifiers.None, explorer, false, out _),
                    "Global recovery owns the matching key release.");
                controller.Group = null;
                Check.That(!controller.Handle(group.Key, true, group.Modifiers, explorer, false, out var disabled) && disabled == null,
                    "Disabling the global shortcut must let it pass through.");
            }
        }
        return Task.CompletedTask;
    }
    private static Task GlobalRelease()
    {
        var controller = Controller();
        Check.That(controller.Handle('E', true, DefaultModifiers, false, false, out var first) && first == SessionAction.RestoreGroup,
            "The global shortcut starts outside Explorer.");
        Check.That(controller.Handle('E', true, DefaultModifiers, true, false, out var repeat) && repeat == null,
            "Moving foreground to Explorer while holding the chord must not restore twice.");
        controller.Group = null;
        Check.That(controller.Handle('E', true, ShortcutModifiers.None, false, false, out var disabledRepeat) && disabledRepeat == null,
            "An already consumed repeat remains owned after disabling the shortcut.");
        Check.That(controller.Handle('E', false, ShortcutModifiers.None, false, false, out _),
            "The global key release remains owned after focus and configuration changes.");
        Check.That(!controller.Handle('E', true, DefaultModifiers, false, false, out _),
            "The next press after disabling must not be intercepted.");
        return Task.CompletedTask;
    }
    private static Task ShortcutEntry()
    {
        var controller = Controller();
        Check.That(!controller.Handle('E', true, DefaultModifiers, false, false, out var action, suppressCommands: true) && action == null,
            "Entering the currently saved group chord in the settings field must not execute recovery.");
        Check.That(!controller.Handle('E', true, DefaultModifiers, false, false, out var repeat) && repeat == null,
            "Losing field focus while the same key remains held must not turn an entry into a command.");
        Check.That(!controller.Handle('E', false, ShortcutModifiers.None, false, false, out _),
            "A chord entered in the field leaves its release available to the field too.");
        Check.That(controller.Handle('E', true, DefaultModifiers, false, false, out var next) && next == SessionAction.RestoreGroup,
            "A later fresh press outside the field must work globally.");
        Check.That(controller.Handle('E', false, ShortcutModifiers.None, false, false, out _, suppressCommands: true),
            "A consumed release is still owned if focus enters a shortcut field afterward.");
        return Task.CompletedTask;
    }
    private static async Task QueuedDelivery(bool dispose)
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(async () =>
        {
            var commands = new List<SessionAction>();
            using var hook = new ExplorerSessionShortcutHook(commands.Add);
            // No native input is sent: exercise the real dispatcher delivery with no surviving origin.
            hook.QueueCommand(SessionAction.RestoreGroup, 0);
            hook.QueueCommand(SessionAction.ReopenTab, 0);
            if (dispose) hook.Dispose();
            await Dispatcher.CurrentDispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
            Check.Equal(dispose ? 0 : 1, commands.Count,
                "Only live global recovery can execute without its originating window.");
            if (!dispose) Check.Equal(SessionAction.RestoreGroup, commands[0], "Tab recovery must retain its foreground guard.");
        }, CancellationToken.None, TaskCreationOptions.None, scheduler).Unwrap();
    }
    private static Task Release()
    {
        var controller = Controller();
        controller.Handle('E', true, DefaultModifiers, true, false, out _);
        controller.Group = null;
        Check.That(controller.Handle('E', false, ShortcutModifiers.None, false, false, out var action) && action == null,
            "A consumed down owns its up even after configuration/focus changes.");
        Check.That(!controller.Handle('E', true, DefaultModifiers, true, false, out _), "Disabled shortcuts pass through.");
        return Task.CompletedTask;
    }
    private static Task NonMatching()
    {
        var controller = Controller();
        foreach (var key in new[] { 'E', 'W' })
        {
            Check.That(!controller.Handle(key, true, DefaultModifiers, true, true, out var injected) && injected == null,
                "Injected keys never trigger either recovery action.");
            foreach (var extra in new[] { ShortcutModifiers.Control, ShortcutModifiers.Shift, ShortcutModifiers.Windows })
            {
                Check.That(!controller.Handle(key, true, DefaultModifiers | extra, false, false, out var action) && action == null,
                    "Extra modifiers must not trigger global recovery.");
                controller.Handle(key, false, ShortcutModifiers.None, false, false, out _);
                Check.That(!controller.Handle(key, true, DefaultModifiers | extra, true, false, out var explorer) && explorer == null,
                    "Extra modifiers must not trigger Explorer recovery either.");
                controller.Handle(key, false, ShortcutModifiers.None, true, false, out _);
            }
        }
        Check.That(!controller.Handle('X', true, DefaultModifiers, true, false, out _), "Unrelated shortcuts pass through.");
        return Task.CompletedTask;
    }
    private static Task Defaults()
    {
        var settings = new AppSettings();
        Check.That(!settings.RestoreTabs && settings.ReopenClosedTab, "Capture/reopen defaults must not enable automatic restoration.");
        Check.That(settings.RestoreGroupShortcutEnabled && settings.ReopenTabShortcutEnabled, "Both optional shortcuts default on.");
        Check.Equal("Alt+E", settings.RestoreGroupShortcut, "The group shortcut defaults to Alt+E.");
        Check.Equal("Alt+W", settings.ReopenTabShortcut, "The individual shortcut defaults to Alt+W.");
        return Task.CompletedTask;
    }
    private static Task MissingShortcutDefaults()
    {
        var directory = Path.Combine(Path.GetTempPath(), "WinTab.Tests", Guid.NewGuid().ToString("N"));
        var path = Path.Combine(directory, "settings.json");
        try
        {
            Directory.CreateDirectory(directory);
            File.WriteAllText(path, """{"WindowHook":false,"Theme":"Dark"}""");
            using var store = new SettingsStore(path);
            Check.That(store.LastError == null, "A configuration without shortcut fields remains valid.");
            Check.Equal("Alt+E", store.Snapshot.RestoreGroupShortcut, "Missing group shortcuts use the new default.");
            Check.Equal("Alt+W", store.Snapshot.ReopenTabShortcut, "Missing tab shortcuts use the new default.");
            Check.That(!store.Snapshot.WindowHook && store.Snapshot.Theme == "Dark", "Other existing preferences remain unchanged.");
        }
        finally { TestCleanup.DeleteDirectory(directory); }
        return Task.CompletedTask;
    }
    private static async Task DefaultPersistence()
    {
        var directory = Path.Combine(Path.GetTempPath(), "WinTab.Tests", Guid.NewGuid().ToString("N"));
        var path = Path.Combine(directory, "settings.json");
        try
        {
            using (var store = new SettingsStore(path))
            {
                Check.That(await store.FlushAsync(), "Fresh defaults must be saved successfully.");
            }
            using var loaded = new SettingsStore(path);
            Check.Equal("Alt+E", loaded.Snapshot.RestoreGroupShortcut, "The new group default survives restart.");
            Check.Equal("Alt+W", loaded.Snapshot.ReopenTabShortcut, "The new tab default survives restart.");
            Check.That(loaded.Snapshot.RestoreGroupShortcutEnabled && loaded.Snapshot.ReopenTabShortcutEnabled,
                "Both new default shortcuts remain enabled after restart.");
        }
        finally { TestCleanup.DeleteDirectory(directory); }
    }
    private static async Task Persistence()
    {
        var directory = Path.Combine(Path.GetTempPath(), "WinTab.Tests", Guid.NewGuid().ToString("N"));
        var path = Path.Combine(directory, "settings.json");
        try
        {
            var store = new SettingsStore(path);
            store.Update(settings => settings with { ReopenClosedTab = false, RestoreGroupShortcutEnabled = false,
                RestoreGroupShortcut = "Ctrl+Alt+F4", ReopenTabShortcutEnabled = false, ReopenTabShortcut = "Alt+7" });
            Check.That(await store.FlushAsync(), "The new fields must be saved.");
            var loaded = new SettingsStore(path).Snapshot;
            Check.That(!loaded.ReopenClosedTab && !loaded.RestoreGroupShortcutEnabled && !loaded.ReopenTabShortcutEnabled,
                "All opt-outs survive a restart.");
            Check.Equal("Ctrl+Alt+F4", loaded.RestoreGroupShortcut, "Custom group chord survives.");
            Check.Equal("Alt+7", loaded.ReopenTabShortcut, "Custom tab chord survives.");
        }
        finally { TestCleanup.DeleteDirectory(directory); }
    }
}
