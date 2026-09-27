using System.Windows.Controls;
using System.Windows.Input;
using WinTab.Hooks;

namespace WinTab.UI.Views.Controls;

/// <summary>An editable shortcut field that captures physical chords directly while focused.</summary>
public sealed class ShortcutTextBox : TextBox
{
    private Key _capturedKey;
    private string? _savedShortcut;
    private bool _capturedKeyReleased;

    public ShortcutTextBox()
    {
        // Shortcut names are Latin key names, not IME text, in both entry modes.
        InputMethod.SetIsInputMethodEnabled(this, false);
        IsVisibleChanged += (_, _) =>
        {
            if (!IsVisible) ResetCapture();
        };
    }

    /// <summary>Unrelated settings refreshes must not discard a draft after this field loses focus.</summary>
    internal void SyncSavedShortcut(string shortcut)
    {
        if (_savedShortcut == shortcut) return;
        _savedShortcut = shortcut;
        Text = shortcut;
    }

    private void ResetCapture()
    {
        _capturedKey = Key.None;
        _capturedKeyReleased = false;
    }

    protected override void OnPreviewKeyDown(KeyEventArgs e)
    {
        var key = ActualKey(e);
        var modifiers = GetModifiers(e.KeyboardDevice);
        if (IsReadOnly || (key == Key.Tab && (modifiers & ~ModifierKeys.Shift) == ModifierKeys.None))
        {
            ResetCapture();
            base.OnPreviewKeyDown(e);
            return;
        }

        if (_capturedKey != Key.None)
        {
            // Own repeats until the whole chord is released, even when modifiers are released first.
            e.Handled = true;
            if (key == _capturedKey) _capturedKeyReleased = false;
            return;
        }
        if (TryGetShortcut(key, modifiers, out var shortcut))
        {
            // Capture editing chords and system keys too, without executing their normal commands.
            e.Handled = true;
            _capturedKey = key;
            if (e.IsRepeat) return;
            Text = shortcut.ToString();
            SelectAll();
            return;
        }

        base.OnPreviewKeyDown(e);
    }

    protected override void OnPreviewKeyUp(KeyEventArgs e)
    {
        if (_capturedKey == Key.None)
        {
            base.OnPreviewKeyUp(e);
            return;
        }

        e.Handled = true;
        if (ActualKey(e) == _capturedKey) _capturedKeyReleased = true;
        if (_capturedKeyReleased && GetModifiers(e.KeyboardDevice) == ModifierKeys.None)
            ResetCapture();
    }

    protected override void OnLostKeyboardFocus(KeyboardFocusChangedEventArgs e)
    {
        ResetCapture();
        base.OnLostKeyboardFocus(e);
    }

    private static Key ActualKey(KeyEventArgs e) => e.Key switch
    {
        Key.System => e.SystemKey,
        Key.ImeProcessed => e.ImeProcessedKey,
        _ => e.Key
    };

    private static ModifierKeys GetModifiers(KeyboardDevice keyboard)
    {
        // WPF's Modifiers omits the Windows keys, so check them explicitly rather than recording a
        // different chord after silently dropping Win.
        var modifiers = keyboard.Modifiers;
        if (keyboard.IsKeyDown(Key.LWin) || keyboard.IsKeyDown(Key.RWin)) modifiers |= ModifierKeys.Windows;
        return modifiers;
    }

    internal static bool TryGetShortcut(Key key, ModifierKeys modifiers, out ExplorerShortcut shortcut)
    {
        shortcut = default;
        const ModifierKeys supported = ModifierKeys.Control | ModifierKeys.Shift | ModifierKeys.Alt;
        if ((modifiers & ~supported) != 0) return false;
        var mapped = (modifiers.HasFlag(ModifierKeys.Control) ? ShortcutModifiers.Control : ShortcutModifiers.None) |
            (modifiers.HasFlag(ModifierKeys.Shift) ? ShortcutModifiers.Shift : ShortcutModifiers.None) |
            (modifiers.HasFlag(ModifierKeys.Alt) ? ShortcutModifiers.Alt : ShortcutModifiers.None);
        return ExplorerShortcut.TryCreate(mapped, KeyInterop.VirtualKeyFromKey(key), out shortcut);
    }
}
