using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.UI.Views.Controls;

internal static class ShortcutTextBoxTests
{
    private const ModifierKeys ChordModifiers = ModifierKeys.Control | ModifierKeys.Shift;
    private const string OriginalText = "shift + control + e";

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("shortcut drafts survive unrelated refreshes and accept newly saved values", () => OnSta(SavedShortcutRefresh));
        yield return ("shortcut key capture shares validation and canonical names with typed input", KeyMapping);
        yield return ("shortcut key capture rejects unsupported keys and modifiers", InvalidKeyMapping);
        yield return ("shortcut key capture works immediately without activating a mode", () => OnSta(RecordChord));
        yield return ("shortcut key capture handles Alt system keys and IME-processed keys", () => OnSta(SystemAndImeKeys));
        yield return ("shortcut key capture leaves incomplete and unsupported chords unchanged", () => OnSta(IncompleteChords));
        yield return ("shortcut key capture owns repeats when modifiers are released first", () => OnSta(ReleaseOrder));
        yield return ("shortcut key capture waits for a repressed main key to be released", () => OnSta(RepressedMainKey));
        yield return ("shortcut key capture preserves text when focus leaves and remains ready on return", () => OnSta(FocusAndNavigation));
        yield return ("shortcut key capture retains ordinary manual editing", () => OnSta(ManualEditing));
        yield return ("shortcut key capture treats editing chords as assignable shortcuts", () => OnSta(EditingChords));
        yield return ("shortcut key capture respects a read-only field", () => OnSta(ReadOnly));
        yield return ("shortcut key capture does not assign a chord already repeating on focus", () => OnSta(InitialRepeat));
        yield return ("shortcut key capture produces a shortcut the Explorer dispatcher can execute", () => OnSta(DispatchRecordedChord));
    }

    private static void SavedShortcutRefresh()
    {
        var input = new ShortcutTextBox();
        input.SyncSavedShortcut("Alt+E");
        Check.Equal("Alt+E", input.Text, "Initialize the field from saved settings.");
        input.Text = "Ctrl+Shift+G";
        Check.That(!input.IsKeyboardFocusWithin, "Exercise a draft after focus left the field.");
        input.SyncSavedShortcut("Alt+E");
        Check.Equal("Ctrl+Shift+G", input.Text, "A refresh of unchanged settings must not erase the draft.");
        input.SyncSavedShortcut("Ctrl+Shift+G");
        input.SyncSavedShortcut("Alt+F2");
        Check.Equal("Alt+F2", input.Text, "A newly committed shortcut must update the displayed value.");
    }

    private static Task KeyMapping()
    {
        foreach (var modifiers in new[] { ModifierKeys.Control, ModifierKeys.Alt, ChordModifiers,
            ModifierKeys.Control | ModifierKeys.Alt, ModifierKeys.Alt | ModifierKeys.Shift,
            ChordModifiers | ModifierKeys.Alt })
        {
            for (var code = 0; code <= 0xFF; code++)
            {
                if (code is not (>= 'A' and <= 'Z' or >= '0' and <= '9' or >= 0x70 and <= 0x87)) continue;
                var key = KeyInterop.KeyFromVirtualKey(code);
                Check.That(ShortcutTextBox.TryGetShortcut(key, modifiers, out var shortcut), "A supported physical chord was rejected.");
                Check.Equal(code, shortcut.Key, "The recorded key must keep its actual virtual-key identity.");
                Check.That(ExplorerShortcut.TryParse(shortcut.ToString(), out var parsed), "Recorded text must pass the existing save validation.");
                Check.Equal(shortcut, parsed, "Physical recording and text entry must describe the same shortcut.");
            }
        }
        Check.That(ShortcutTextBox.TryGetShortcut(Key.D7, ChordModifiers | ModifierKeys.Alt, out var digit), "Shifted top-row digits are supported.");
        Check.Equal("Ctrl+Shift+Alt+7", digit.ToString(), "Record the physical digit, not the shifted punctuation it would type.");
        return Task.CompletedTask;
    }

    private static Task InvalidKeyMapping()
    {
        foreach (var modifiers in new[] { ModifierKeys.None, ModifierKeys.Shift, ModifierKeys.Windows,
            ModifierKeys.Control | ModifierKeys.Windows, ChordModifiers | ModifierKeys.Windows, (ModifierKeys)16 })
            Check.That(!ShortcutTextBox.TryGetShortcut(Key.E, modifiers, out _), "Unsupported modifiers must not be silently dropped.");
        foreach (var key in new[] { Key.None, Key.LeftCtrl, Key.RightCtrl, Key.LeftShift, Key.RightShift,
            Key.LeftAlt, Key.RightAlt, Key.LWin, Key.RWin, Key.Escape, Key.Tab, Key.Back, Key.Delete,
            Key.Space, Key.OemPlus, Key.NumPad0, Key.NumPad7, Key.NumPad9, Key.System, Key.ImeProcessed })
            Check.That(!ShortcutTextBox.TryGetShortcut(key, ChordModifiers, out _), "Unsupported keys must not turn into unrelated letter shortcuts: " + key);
        Check.That(!ExplorerShortcut.TryCreate((ShortcutModifiers)8 | ShortcutModifiers.Control, 'E', out _),
            "The common validator must reject unknown modifier bits as well.");
        return Task.CompletedTask;
    }

    private static void RecordChord()
    {
        var input = NewInput();
        var keyboard = new TestKeyboard();
        Check.That(!input.IsReadOnly, "The field is editable and ready without any activation step.");
        Down(input, keyboard, Key.RightCtrl, ModifierKeys.Control);
        Check.Equal(OriginalText, input.Text, "Modifier-only input must not replace the previous shortcut.");
        Down(input, keyboard, Key.RightShift, ChordModifiers);
        Check.That(Down(input, keyboard, Key.T, ChordModifiers).Handled, "The chord must not reach normal text editing.");
        Check.Equal("Ctrl+Shift+T", input.Text, "The physical chord fills the input immediately.");
        Check.Equal(input.Text.Length, input.SelectionLength, "The whole shortcut is selected for replacement or manual editing.");
        Check.That(Up(input, keyboard, Key.T, ChordModifiers).Handled, "The main-key release is owned.");
        Up(input, keyboard, Key.RightShift, ModifierKeys.Control);
        Check.That(Up(input, keyboard, Key.RightCtrl, ModifierKeys.None).Handled, "The last release is consumed too.");
        Check.That(!input.IsReadOnly, "Capturing a chord never switches to read-only mode.");
        Down(input, keyboard, Key.E, ChordModifiers);
        Check.Equal("Ctrl+Shift+E", input.Text, "Another chord can immediately replace the first one without a button.");
        Up(input, keyboard, Key.E, ModifierKeys.None);
    }

    private static void SystemAndImeKeys()
    {
        var input = NewInput();
        var keyboard = new TestKeyboard();
        var alt = Down(input, keyboard, Key.F4, ModifierKeys.Alt, "MarkSystem");
        Check.Equal(Key.System, alt.Key, "Exercise WPF's actual system-key representation.");
        Check.That(alt.Handled, "Alt+F4 must be captured rather than closing the settings window.");
        Check.Equal("Alt+F4", input.Text, "Use SystemKey instead of the Key.System placeholder.");
        Up(input, keyboard, Key.F4, ModifierKeys.Alt, "MarkSystem");
        Up(input, keyboard, Key.LeftAlt, ModifierKeys.None, "MarkSystem");

        var ime = Down(input, keyboard, Key.E, ChordModifiers, "MarkImeProcessed");
        Check.Equal(Key.ImeProcessed, ime.Key, "Exercise WPF's IME-processed representation.");
        Check.Equal("Ctrl+Shift+E", input.Text, "An IME-processed event must use the underlying physical key.");
        Check.That(!InputMethod.GetIsInputMethodEnabled(input), "IME text composition is not needed for shortcut names.");
        Up(input, keyboard, Key.E, ModifierKeys.None, "MarkImeProcessed");
    }

    private static void IncompleteChords()
    {
        var input = NewInput();
        var keyboard = new TestKeyboard();
        foreach (var (key, modifiers) in new[] { (Key.LeftCtrl, ModifierKeys.Control),
            (Key.NumPad7, ChordModifiers), (Key.OemPlus, ModifierKeys.Control),
            (Key.E, ModifierKeys.Control | ModifierKeys.Windows) })
        {
            Down(input, keyboard, key, modifiers);
            Up(input, keyboard, key, ModifierKeys.None);
            Check.Equal(OriginalText, input.Text, "An unsupported chord must not be assigned as a different shortcut.");
        }
        Down(input, keyboard, Key.T, ChordModifiers);
        Check.Equal("Ctrl+Shift+T", input.Text, "A valid fresh chord still works after unsupported input.");
        Up(input, keyboard, Key.T, ModifierKeys.None);
    }

    private static void ReleaseOrder()
    {
        var input = NewInput();
        var keyboard = new TestKeyboard();
        Down(input, keyboard, Key.T, ChordModifiers);
        Up(input, keyboard, Key.LeftCtrl, ModifierKeys.Shift);
        Up(input, keyboard, Key.LeftShift, ModifierKeys.None);
        Check.That(Down(input, keyboard, Key.T, ModifierKeys.None, repeat: true).Handled, "A now-unmodified repeat is still consumed.");
        Check.Equal("Ctrl+Shift+T", input.Text, "Repeats must not append letters or replace the chord.");
        Up(input, keyboard, Key.T, ModifierKeys.None);
        Check.That(!Down(input, keyboard, Key.E, ModifierKeys.None).Handled, "The main key's release makes ordinary typing available again.");
    }

    private static void RepressedMainKey()
    {
        var input = NewInput();
        var keyboard = new TestKeyboard();
        Down(input, keyboard, Key.T, ChordModifiers);
        Up(input, keyboard, Key.T, ChordModifiers);
        Down(input, keyboard, Key.T, ChordModifiers);
        Up(input, keyboard, Key.LeftCtrl, ModifierKeys.Shift);
        Up(input, keyboard, Key.LeftShift, ModifierKeys.None);
        Check.That(Down(input, keyboard, Key.T, ModifierKeys.None, repeat: true).Handled, "A repressed letter remains owned until it is released again.");
        Up(input, keyboard, Key.T, ModifierKeys.None);
        Check.Equal("Ctrl+Shift+T", input.Text, "Only the first complete chord is captured while keys remain held.");
        Down(input, keyboard, Key.E, ChordModifiers);
        Check.Equal("Ctrl+Shift+E", input.Text, "A new complete chord can be captured after the last release.");
        Up(input, keyboard, Key.E, ModifierKeys.None);
    }

    private static void FocusAndNavigation()
    {
        foreach (var modifiers in new[] { ModifierKeys.None, ModifierKeys.Shift })
        {
            var input = NewInput();
            var keyboard = new TestKeyboard();
            Down(input, keyboard, Key.T, ChordModifiers);
            Check.That(!Down(input, keyboard, Key.Tab, modifiers).Handled, "Tab and Shift+Tab retain normal navigation.");
            Check.Equal("Ctrl+Shift+T", input.Text, "Navigation does not undo the captured shortcut.");
            Down(input, keyboard, Key.E, ChordModifiers);
            Check.Equal("Ctrl+Shift+E", input.Text, "Navigation clears key ownership without disabling capture.");
        }

        var focusedInput = NewInput();
        var focusedKeyboard = new TestKeyboard();
        Down(focusedInput, focusedKeyboard, Key.T, ChordModifiers);
        focusedInput.RaiseEvent(new KeyboardFocusChangedEventArgs(focusedKeyboard, 0, focusedInput, new TextBox())
        {
            RoutedEvent = Keyboard.LostKeyboardFocusEvent
        });
        Check.Equal("Ctrl+Shift+T", focusedInput.Text, "Leaving the field preserves its captured value.");
        Check.That(!Down(focusedInput, focusedKeyboard, Key.E, ModifierKeys.None).Handled, "Focus loss clears an incomplete release sequence.");
        Down(focusedInput, focusedKeyboard, Key.E, ChordModifiers);
        Check.Equal("Ctrl+Shift+E", focusedInput.Text, "Returning to the field needs no recording-mode activation.");
    }

    private static void ManualEditing()
    {
        var input = NewInput();
        var keyboard = new TestKeyboard();
        AssertEditingKeysPassThrough(input, keyboard);
        Down(input, keyboard, Key.T, ChordModifiers);
        Up(input, keyboard, Key.T, ModifierKeys.None);
        AssertEditingKeysPassThrough(input, keyboard);
    }

    private static void AssertEditingKeysPassThrough(ShortcutTextBox input, TestKeyboard keyboard)
    {
        Check.That(!input.IsReadOnly, "The input always remains manually editable.");
        foreach (var key in new[] { Key.C, Key.Back, Key.Delete, Key.Left, Key.Right, Key.Tab, Key.Escape, Key.OemPlus })
            Check.That(!Down(input, keyboard, key, ModifierKeys.None).Handled, "Ordinary typing and navigation must pass through: " + key);
        Check.That(!Down(input, keyboard, Key.E, ModifierKeys.Shift).Handled, "Shifted letters remain ordinary text input.");
    }

    private static void EditingChords()
    {
        var input = NewInput();
        var keyboard = new TestKeyboard();
        foreach (var key in new[] { Key.A, Key.C, Key.V, Key.X, Key.Z, Key.Y })
        {
            Check.That(Down(input, keyboard, key, ModifierKeys.Control).Handled, "Editing commands must not bypass automatic capture: Ctrl+" + key);
            Check.Equal("Ctrl+" + key, input.Text, "Editing chords are assignable just like every supported chord.");
            Up(input, keyboard, key, ModifierKeys.None);
        }
    }

    private static void ReadOnly()
    {
        var input = NewInput();
        var keyboard = new TestKeyboard();
        input.IsReadOnly = true;
        Down(input, keyboard, Key.T, ChordModifiers);
        Check.Equal(OriginalText, input.Text, "Capture must not override a read-only field.");
        input.IsReadOnly = false;
        Down(input, keyboard, Key.T, ChordModifiers);
        Check.Equal("Ctrl+Shift+T", input.Text, "Making the field editable restores capture automatically.");
    }

    private static void InitialRepeat()
    {
        var input = NewInput();
        var keyboard = new TestKeyboard();
        Check.That(Down(input, keyboard, Key.T, ChordModifiers, repeat: true).Handled, "A held chord must not run a text command on focus.");
        Check.Equal(OriginalText, input.Text, "An already-repeating key is not a fresh assignment.");
        Up(input, keyboard, Key.LeftCtrl, ModifierKeys.Shift);
        Up(input, keyboard, Key.LeftShift, ModifierKeys.None);
        Check.That(Down(input, keyboard, Key.T, ModifierKeys.None, repeat: true).Handled, "Repeats stay owned after modifiers are released.");
        Up(input, keyboard, Key.T, ModifierKeys.None);
        Down(input, keyboard, Key.T, ChordModifiers);
        Check.Equal("Ctrl+Shift+T", input.Text, "A fresh press after the held key is released can assign the shortcut.");
    }

    private static void DispatchRecordedChord()
    {
        var input = NewInput();
        var otherInput = NewInput();
        var keyboard = new TestKeyboard();
        Down(input, keyboard, Key.F8, ModifierKeys.Control | ModifierKeys.Alt, "MarkSystem");
        Up(input, keyboard, Key.F8, ModifierKeys.None);
        Check.Equal(OriginalText, otherInput.Text, "Capturing one recovery action must not edit the other field.");
        Check.That(ExplorerShortcut.TryParse(input.Text, out var saved), "Captured text is accepted by save validation.");
        var dispatch = new ExplorerShortcutDispatch { Tab = saved };
        var modifiers = ShortcutModifiers.Control | ShortcutModifiers.Alt;
        Check.That(!dispatch.Handle(0x77, true, modifiers, false, false, out _), "Captured shortcuts remain Explorer-scoped.");
        dispatch.Handle(0x77, false, modifiers, false, false, out _);
        Check.That(dispatch.Handle(0x77, true, modifiers, true, false, out var action) && action == SessionAction.ReopenTab,
            "The exact physical chord captured by the UI must execute through the existing dispatcher.");
    }

    private static ShortcutTextBox NewInput() => new() { Text = OriginalText };

    private static async Task OnSta(Action body)
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(body, CancellationToken.None, TaskCreationOptions.None, scheduler);
    }

    private static KeyEventArgs Down(ShortcutTextBox input, TestKeyboard keyboard, Key key, ModifierKeys modifiers,
        string? marker = null, bool repeat = false) => SendKey(input, keyboard, key, modifiers, Keyboard.PreviewKeyDownEvent, marker, repeat);

    private static KeyEventArgs Up(ShortcutTextBox input, TestKeyboard keyboard, Key key, ModifierKeys modifiers,
        string? marker = null) => SendKey(input, keyboard, key, modifiers, Keyboard.PreviewKeyUpEvent, marker, false);

    private static KeyEventArgs SendKey(ShortcutTextBox input, TestKeyboard keyboard, Key key, ModifierKeys modifiers,
        RoutedEvent routedEvent, string? marker, bool repeat)
    {
        keyboard.PressedModifiers = modifiers;
        var args = new KeyEventArgs(keyboard, keyboard.InputSource, 0, key) { RoutedEvent = routedEvent };
        // WPF's input provider marks these internally; use the same event representation without sending
        // any physical input to the user's desktop or depending on their current keyboard state.
        if (marker != null)
            typeof(KeyEventArgs).GetMethod(marker, BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(args, null);
        if (repeat)
            typeof(KeyEventArgs).GetMethod("SetRepeat", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(args, [true]);
        input.RaiseEvent(args);
        return args;
    }

    private sealed class TestInputSource : PresentationSource
    {
        public override Visual RootVisual { get; set; } = null!;
        public override bool IsDisposed => false;
        protected override CompositionTarget GetCompositionTargetCore() => null!;
    }

    private sealed class TestKeyboard() : KeyboardDevice(InputManager.Current)
    {
        public PresentationSource InputSource { get; } = new TestInputSource();
        public ModifierKeys PressedModifiers { get; set; }
        protected override KeyStates GetKeyStatesFromSystem(Key key)
        {
            var modifier = key switch
            {
                Key.LeftCtrl or Key.RightCtrl => ModifierKeys.Control,
                Key.LeftShift or Key.RightShift => ModifierKeys.Shift,
                Key.LeftAlt or Key.RightAlt => ModifierKeys.Alt,
                Key.LWin or Key.RWin => ModifierKeys.Windows,
                _ => ModifierKeys.None
            };
            return (PressedModifiers & modifier) != 0 ? KeyStates.Down : KeyStates.None;
        }
    }
}
