using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace WinTab.Hooks;

[Flags]
internal enum ShortcutModifiers { None = 0, Control = 1, Shift = 2, Alt = 4 }
internal enum SessionAction { RestoreGroup, ReopenTab }

internal readonly record struct ExplorerShortcut(ShortcutModifiers Modifiers, int Key)
{
    public static bool TryParse(string? text, out ExplorerShortcut shortcut)
    {
        shortcut = default;
        if (string.IsNullOrWhiteSpace(text)) return false;
        var parts = text.Split('+', StringSplitOptions.TrimEntries);
        if (parts.Length < 2 || parts.Length > 4) return false;
        var modifiers = ShortcutModifiers.None;
        foreach (var part in parts[..^1])
        {
            var modifier = part.ToUpperInvariant() switch
            {
                "CTRL" or "CONTROL" => ShortcutModifiers.Control,
                "SHIFT" => ShortcutModifiers.Shift,
                "ALT" => ShortcutModifiers.Alt,
                _ => ShortcutModifiers.None
            };
            if (modifier == ShortcutModifiers.None || modifiers.HasFlag(modifier)) return false;
            modifiers |= modifier;
        }
        if ((modifiers & (ShortcutModifiers.Control | ShortcutModifiers.Alt)) == 0) return false;
        var key = parts[^1].ToUpperInvariant();
        var code = key.Length == 1 && (key[0] is >= 'A' and <= 'Z' or >= '0' and <= '9') ? key[0] : 0;
        if (code == 0 && key.StartsWith('F') && int.TryParse(key.AsSpan(1), NumberStyles.None,
            CultureInfo.InvariantCulture, out var number) && number is >= 1 and <= 24)
            code = 0x70 + number - 1;
        if (code == 0) return false;
        shortcut = new ExplorerShortcut(modifiers, code);
        return true;
    }

    public override string ToString() => string.Join("+", new[]
    {
        Modifiers.HasFlag(ShortcutModifiers.Control) ? "Ctrl" : null,
        Modifiers.HasFlag(ShortcutModifiers.Shift) ? "Shift" : null,
        Modifiers.HasFlag(ShortcutModifiers.Alt) ? "Alt" : null,
        Key is >= 0x70 and <= 0x87 ? "F" + (Key - 0x70 + 1) : ((char)Key).ToString()
    }.Where(part => part != null));
}

/// <summary>Only Explorer-scoped presses are consumed. A consumed down owns its repeats and matching up.</summary>
internal sealed class ExplorerShortcutDispatch
{
    private readonly HashSet<int> _pressed = [];
    private readonly HashSet<int> _consumed = [];
    public ExplorerShortcut? Group { get; set; }
    public ExplorerShortcut? Tab { get; set; }

    public void Reset()
    {
        _pressed.Clear();
        _consumed.Clear();
    }

    public bool Handle(int key, bool down, ShortcutModifiers modifiers, bool explorer, bool injected,
        out SessionAction? action)
    {
        action = null;
        if (injected) return false;
        if (!down)
        {
            _pressed.Remove(key);
            return _consumed.Remove(key);
        }
        if (!_pressed.Add(key)) return _consumed.Contains(key);
        if (!explorer) return false;
        var stroke = new ExplorerShortcut(modifiers, key);
        action = Group == stroke ? SessionAction.RestoreGroup : Tab == stroke ? SessionAction.ReopenTab : null;
        if (action == null) return false;
        _consumed.Add(key);
        return true;
    }
}
