using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using WinTab.Hooks;
using WinTab.Managers;

internal static class ExplorerShortcutTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("session shortcuts normalize and validate configurable chords", Parsing);
        yield return ("session shortcuts leave browsers alone and suppress autorepeat", Dispatch);
        yield return ("session shortcuts preserve consumed key-up across focus changes", Release);
        yield return ("session shortcuts ignore injected input and extra modifiers", NonMatching);
        yield return ("session recovery settings default on without enabling automatic restore", Defaults);
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
        ExplorerShortcut.TryParse("Ctrl+Shift+E", out var group);
        ExplorerShortcut.TryParse("Ctrl+Shift+T", out var tab);
        return new ExplorerShortcutDispatch { Group = group, Tab = tab };
    }
    private const ShortcutModifiers DefaultModifiers = ShortcutModifiers.Control | ShortcutModifiers.Shift;
    private static Task Dispatch()
    {
        var controller = Controller();
        Check.That(!controller.Handle('T', true, DefaultModifiers, false, false, out var browser) && browser == null,
            "Browser Ctrl+Shift+T must pass through.");
        controller.Handle('T', false, DefaultModifiers, false, false, out _);
        Check.That(controller.Handle('T', true, DefaultModifiers, true, false, out var first) && first == SessionAction.ReopenTab,
            "Explorer invokes the command once.");
        Check.That(controller.Handle('T', true, DefaultModifiers, true, false, out var repeat) && repeat == null,
            "Holding the shortcut does not reopen the whole history.");
        Check.That(controller.Handle('T', false, DefaultModifiers, true, false, out _), "Its key-up is consumed.");
        Check.That(controller.Handle('T', true, DefaultModifiers, true, false, out var second) && second == SessionAction.ReopenTab,
            "A later physical press can invoke again.");
        return Task.CompletedTask;
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
        Check.That(!controller.Handle('T', true, DefaultModifiers, true, true, out _), "Injected keys never trigger recovery.");
        Check.That(!controller.Handle('T', true, DefaultModifiers | ShortcutModifiers.Alt, true, false, out _), "Extra modifiers do not match.");
        controller.Handle('T', false, DefaultModifiers, true, false, out _);
        Check.That(!controller.Handle('X', true, DefaultModifiers, true, false, out _), "Unrelated shortcuts pass through.");
        return Task.CompletedTask;
    }
    private static Task Defaults()
    {
        var settings = new AppSettings();
        Check.That(!settings.RestoreTabs && settings.ReopenClosedTab, "Capture/reopen defaults must not enable automatic restoration.");
        Check.That(settings.RestoreGroupShortcutEnabled && settings.ReopenTabShortcutEnabled, "Both optional shortcuts default on.");
        Check.Equal("Ctrl+Shift+E", settings.RestoreGroupShortcut, "The group shortcut uses E.");
        Check.Equal("Ctrl+Shift+T", settings.ReopenTabShortcut, "The individual shortcut uses T.");
        return Task.CompletedTask;
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
