using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using WinTab.Hooks;
using WinTab.Managers;
using WinTab.UI.Desktop;

internal static class DesktopBridgeTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("desktop bridge accepts typed settings and recovery commands", Commands);
        yield return ("desktop bridge rejects unknown commands and malformed values", InvalidCommands);
        yield return ("desktop bridge preserves the shared frontend and Go state contract", StateContract);
        yield return ("desktop bridge reads Unicode frames and detects clean disconnects", Framing);
        yield return ("desktop bridge rejects oversized and incomplete frames", InvalidFrames);
        yield return ("desktop shortcut capture leaves the renderer's physical chord untouched", ShortcutScope);
    }

    private static Task Commands()
    {
        foreach (var key in new[] { "windowHook", "reuseTabs", "restoreTabs", "restoreSingleTab",
            "restoreOnAnyFolder", "reopenClosedTab", "doubleClickCloseTab",
            "doubleClickCloseIncludeNotepad", "middleClickForegroundTab", "wheelSwitchTab",
            "showTrayIcon", "autoUpdate", "startup" })
        {
            foreach (var enabled in new[] { false, true })
            {
                var command = Parse("set", new { key, value = enabled });
                Check.Equal(key, command.Key);
                Check.Equal(enabled, command.Enabled);
            }
        }
        foreach (var (key, value) in new[] { ("language", "zh-CN"), ("language", "en-US"),
            ("theme", "Light"), ("theme", "Dark"), ("wheelSwitchSensitivity", "Low"),
            ("wheelSwitchSensitivity", "Medium"), ("wheelSwitchSensitivity", "High") })
            Check.Equal(value, Parse("set", new { key, value }).Text);

        // Chord normalization remains owned by the existing shared shortcut validator.
        var shortcuts = Parse("shortcuts", new { groupEnabled = false, group = "shift + control + e",
            tabEnabled = true, tab = "Alt+F24" });
        Check.That(!shortcuts.GroupEnabled && shortcuts.TabEnabled, "The two enabled flags must stay independent");
        Check.Equal("shift + control + e", shortcuts.GroupShortcut);
        Check.Equal("Alt+F24", shortcuts.TabShortcut);
        Check.That(Parse("restore", new { group = true }).Enabled, "Group restore is selected explicitly");
        Check.That(!Parse("restore", new { group = false }).Enabled, "Single-tab restore is selected explicitly");
        var size = Parse("size", new { width = 920.5, height = 760 });
        Check.Equal(920.5, size.Size.Width);
        Check.Equal(760.0, size.Size.Height);
        foreach (var method in new[] { "state", "update", "logs" })
        {
            Check.Equal(method, Parse(method, null).Method);
            Check.Equal(17L, Parse(method, null).Id, "Replies must retain the request id");
        }
        return Task.CompletedTask;
    }

    private static Task InvalidCommands()
    {
        foreach (var line in new[] { "null", "{}", "[]", "{",
            "{\"id\":0,\"method\":\"state\"}", "{\"id\":-1,\"method\":\"state\"}",
            "{\"id\":1,\"method\":\"shell\"}", "{\"id\":1,\"method\":\"state\",\"executable\":\"cmd.exe\"}",
            Request("set", new { key = "unknown", value = true }),
            Request("set", new { key = "windowHook", value = "false" }),
            Request("set", new { key = "windowHook", value = 0 }),
            Request("set", new { key = "windowHook" }),
            Request("set", new { key = "language", value = "fr-FR" }),
            Request("set", new { key = "theme", value = "System" }),
            Request("set", new { key = "wheelSwitchSensitivity", value = "2" }),
            Request("shortcuts", new { groupEnabled = true, group = "", tabEnabled = false, tab = "Alt+W" }),
            Request("restore", new { group = "true" }),
            Request("size", new { width = 399, height = 700 }),
            Request("size", new { width = 900, height = 10001 }),
            "{\"id\":1,\"method\":\"size\",\"params\":{\"width\":1e999,\"height\":700}}",
            new string(' ', DesktopCommand.MaximumMessageLength + 1) })
        {
            try { DesktopCommand.Parse(line); }
            catch (Exception error) when (error is FormatException or JsonException) { continue; }
            throw new InvalidOperationException("Invalid desktop command was accepted: " + line[..Math.Min(line.Length, 180)]);
        }
        return Task.CompletedTask;
    }

    private static Task StateContract()
    {
        var state = new DesktopState(1, 1, "2.4.2", new AppSettings { Language = "zh-CN" },
            false, true, true, false, false, null, null, null, null, null);
        using var expected = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "DesktopState.json")));
        var actual = JsonSerializer.SerializeToElement(state, DesktopCommand.Json);
        Check.That(JsonElement.DeepEquals(expected.RootElement, actual),
            "The resident snapshot must match the fixture consumed by Go and the frontend");
        return Task.CompletedTask;
    }

    private static async Task Framing()
    {
        using var input = new StringReader("中文🙂\r\nsecond\n");
        Check.Equal("中文🙂", await DesktopUiProcess.ReadMessageAsync(input));
        Check.Equal("second", await DesktopUiProcess.ReadMessageAsync(input));
        Check.That(await DesktopUiProcess.ReadMessageAsync(input) is null, "EOF between frames is a normal disconnect");
        using var boundary = new StringReader(new string('x', DesktopCommand.MaximumMessageLength) + "\n");
        Check.Equal(DesktopCommand.MaximumMessageLength, (await DesktopUiProcess.ReadMessageAsync(boundary))!.Length);
    }

    private static async Task InvalidFrames()
    {
        foreach (var text in new[] { "unterminated", new string('x', DesktopCommand.MaximumMessageLength + 1) + "\n" })
        {
            using var reader = new StringReader(text);
            try { await DesktopUiProcess.ReadMessageAsync(reader); }
            catch (IOException) { continue; }
            throw new InvalidOperationException("Damaged pipe frame was accepted");
        }
    }

    private static Task ShortcutScope()
    {
        foreach (var (foreground, renderer, wpf, suppressed) in new[] {
            (42u, 42u, false, true), (42u, 91u, false, false), (42u, 0u, false, false),
            (42u, 91u, true, true) })
        {
            ExplorerShortcut.TryParse("Alt+E", out var shortcut);
            var dispatch = new ExplorerShortcutDispatch { Group = shortcut };
            var consumed = dispatch.Handle('E', true, ShortcutModifiers.Alt, false, false, out var action,
                suppressCommands: ExplorerSessionShortcutHook.SuppressForSettings(foreground, renderer, wpf));
            Check.Equal(!suppressed, consumed);
            Check.Equal<SessionAction?>(suppressed ? null : SessionAction.RestoreGroup, action);
        }
        return Task.CompletedTask;
    }

    private static string Request(string method, object? parameters) =>
        JsonSerializer.Serialize(new { id = 17, method, @params = parameters });
    private static DesktopCommand Parse(string method, object? parameters) => DesktopCommand.Parse(Request(method, parameters));
}
