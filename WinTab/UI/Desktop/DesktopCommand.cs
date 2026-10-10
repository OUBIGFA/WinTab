using System;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;

namespace WinTab.UI.Desktop;

/// <summary>The small, versioned allowlist exposed to the bundled UI, never a general RPC or shell endpoint.</summary>
internal sealed record DesktopRequest(long Id, string Method, JsonElement Params);

internal sealed record DesktopCommand(long Id, string Method)
{
    internal string Key { get; init; } = string.Empty;
    internal string Text { get; init; } = string.Empty;
    internal bool Enabled { get; init; }
    internal bool GroupEnabled { get; init; }
    internal bool TabEnabled { get; init; }
    internal string GroupShortcut { get; init; } = string.Empty;
    internal string TabShortcut { get; init; } = string.Empty;
    internal Size Size { get; init; }

    internal const int MaximumMessageLength = 65536;
    internal static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web)
    {
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow
    };

    internal static DesktopCommand Parse(string line)
    {
        if (line.Length > MaximumMessageLength) throw new FormatException("Message too large");
        var request = JsonSerializer.Deserialize<DesktopRequest>(line, Json)
            ?? throw new FormatException("Missing request");
        if (request.Id <= 0) throw new FormatException("Invalid request id");
        var command = new DesktopCommand(request.Id, request.Method);
        var args = request.Params;
        switch (request.Method)
        {
            case "state":
            case "update":
            case "logs":
                return command;
            case "set":
                var key = String(args, "key");
                command = command with { Key = key };
                return key switch
                {
                    "windowHook" or "reuseTabs" or "restoreTabs" or "restoreSingleTab" or
                    "restoreOnAnyFolder" or "reopenClosedTab" or "doubleClickCloseTab" or
                    "doubleClickCloseIncludeNotepad" or "middleClickForegroundTab" or "wheelSwitchTab" or
                    "showTrayIcon" or "autoUpdate" or "startup" => command with { Enabled = Bool(args, "value") },
                    "language" => command with { Text = Choice(args, "value", "zh-CN", "en-US") },
                    "theme" => command with { Text = Choice(args, "value", "Light", "Dark") },
                    "wheelSwitchSensitivity" => command with { Text = Choice(args, "value", "Low", "Medium", "High") },
                    _ => throw new FormatException("Unknown setting")
                };
            case "shortcuts":
                return command with
                {
                    GroupEnabled = Bool(args, "groupEnabled"), TabEnabled = Bool(args, "tabEnabled"),
                    GroupShortcut = String(args, "group"), TabShortcut = String(args, "tab")
                };
            case "restore":
                return command with { Enabled = Bool(args, "group") };
            case "size":
                var width = Number(args, "width");
                var height = Number(args, "height");
                if (width is < 400 or > 10000 || height is < 300 or > 10000)
                    throw new FormatException("Invalid window size");
                return command with { Size = new Size(width, height) };
            default:
                throw new FormatException("Unknown command");
        }
    }

    private static JsonElement Property(JsonElement element, string name) =>
        element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out var value)
            ? value : throw new FormatException("Missing argument: " + name);

    private static string String(JsonElement element, string name)
    {
        var value = Property(element, name);
        return value.ValueKind == JsonValueKind.String && value.GetString() is { Length: > 0 and <= 128 } text
            ? text : throw new FormatException("Invalid string: " + name);
    }

    private static bool Bool(JsonElement element, string name) => Property(element, name).ValueKind switch
    {
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        _ => throw new FormatException("Invalid boolean: " + name)
    };

    private static double Number(JsonElement element, string name)
    {
        var value = Property(element, name);
        return value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number) && double.IsFinite(number)
            ? number : throw new FormatException("Invalid number: " + name);
    }

    private static string Choice(JsonElement element, string name, params string[] values)
    {
        var value = String(element, name);
        return Array.IndexOf(values, value) >= 0 ? value : throw new FormatException("Invalid choice: " + name);
    }
}
