using System.Collections.ObjectModel;
using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WinTab.Core.Enums;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Persistence;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace WinTab.App.ViewModels;

/// <summary>
/// Wraps a <see cref="HotKeyBinding"/> to provide display-friendly properties.
/// </summary>
public partial class HotKeyBindingViewModel : ObservableObject
{
    private readonly HotKeyBinding _binding;

    public HotKeyBindingViewModel(HotKeyBinding binding)
    {
        _binding = binding;
    }

    public HotKeyAction Action => _binding.Action;

    public string ActionDisplayName => _binding.Action switch
    {
        HotKeyAction.NextTab => "Next Tab",
        HotKeyAction.PreviousTab => "Previous Tab",
        HotKeyAction.CloseTab => "Close Tab",
        HotKeyAction.NewInstance => "New Instance",
        HotKeyAction.DetachTab => "Detach Tab",
        HotKeyAction.MoveTabLeft => "Move Tab Left",
        HotKeyAction.MoveTabRight => "Move Tab Right",
        HotKeyAction.ToggleGrouping => "Toggle Grouping",
        _ => _binding.Action.ToString()
    };

    public bool Enabled
    {
        get => _binding.Enabled;
        set
        {
            if (_binding.Enabled == value) return;
            _binding.Enabled = value;
            OnPropertyChanged();
        }
    }

    public string BindingDisplay => string.IsNullOrEmpty(_binding.DisplayString) || _binding.Key == 0
        ? "(none)"
        : _binding.DisplayString;

    public uint Modifiers
    {
        get => _binding.Modifiers;
        set
        {
            _binding.Modifiers = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(BindingDisplay));
        }
    }

    public uint Key
    {
        get => _binding.Key;
        set
        {
            _binding.Key = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(BindingDisplay));
        }
    }

    public HotKeyBinding GetBinding() => _binding;
}

public partial class ShortcutsViewModel : ObservableObject
{
    private readonly AppSettings _settings;
    private readonly SettingsStore _settingsStore;
    private readonly Logger _logger;

    public ObservableCollection<HotKeyBindingViewModel> HotKeys { get; }

    [ObservableProperty]
    private bool _isRecording;

    private HotKeyBindingViewModel? _recordingTarget;

    public ShortcutsViewModel(
        AppSettings settings,
        SettingsStore settingsStore,
        Logger logger)
    {
        _settings = settings;
        _settingsStore = settingsStore;
        _logger = logger;

        // Initialize hotkeys; if settings has none, populate defaults for each action
        if (settings.HotKeys.Count == 0)
        {
            foreach (HotKeyAction action in Enum.GetValues<HotKeyAction>())
            {
                settings.HotKeys.Add(new HotKeyBinding
                {
                    Action = action,
                    Enabled = false,
                    Modifiers = 0,
                    Key = 0
                });
            }
        }

        HotKeys = new ObservableCollection<HotKeyBindingViewModel>(
            settings.HotKeys.Select(hk => new HotKeyBindingViewModel(hk)));
    }

    [RelayCommand]
    private void StartRecording(HotKeyBindingViewModel? binding)
    {
        if (binding is null) return;

        _recordingTarget = binding;
        IsRecording = true;
        _logger.Info($"Recording shortcut for action: {binding.Action}");
    }

    /// <summary>
    /// Called from the page code-behind when a key is pressed during recording.
    /// </summary>
    public void RecordKey(KeyEventArgs e)
    {
        if (!IsRecording || _recordingTarget is null)
            return;

        Key key = e.Key == Key.System ? e.SystemKey : e.Key;

        // Ignore lone modifier keys
        if (key is Key.LeftShift or Key.RightShift or
            Key.LeftCtrl or Key.RightCtrl or
            Key.LeftAlt or Key.RightAlt or
            Key.LWin or Key.RWin)
            return;

        // Build Win32 modifier flags
        uint modifiers = 0;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Alt)) modifiers |= 0x0001;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control)) modifiers |= 0x0002;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)) modifiers |= 0x0004;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Windows)) modifiers |= 0x0008;

        // Convert WPF key to Win32 virtual key code
        uint vk = (uint)KeyInterop.VirtualKeyFromKey(key);

        _recordingTarget.Modifiers = modifiers;
        _recordingTarget.Key = vk;

        _logger.Info($"Recorded shortcut for {_recordingTarget.Action}: {_recordingTarget.BindingDisplay}");

        // Stop recording
        IsRecording = false;
        _recordingTarget = null;

        // Sync back to settings
        SyncToSettings();
    }

    private void SyncToSettings()
    {
        _settings.HotKeys = HotKeys.Select(vm => vm.GetBinding()).ToList();
        _settingsStore.SaveDebounced(_settings);
    }
}
