using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using WinTab.Hooks;
using WinTab.Managers;
using WinTab.UI.Localization;
using WinTab.UI.Views.Controls;

namespace WinTab.UI.Desktop;

internal interface IApplicationUi
{
    void ShowMainWindow();
    Task OpenRecycleBinAsync();
}

/// <summary>A complete snapshot, so pipe replies and tray-originated changes never leave a half-updated form.</summary>
internal sealed record DesktopState(int ProtocolVersion, long Revision, string Version, AppSettings Settings,
    bool Startup, bool ShellReady, bool RecordClosedTabs, bool SessionBusy, bool UpdateBusy,
    string? HookError, string? ShortcutError, string? SessionFeedback, string? UpdateFeedback, string? StorageError);

/// <summary>
/// Resident application for x64/arm64. The renderer may come and go, but hooks, recording, tray,
/// settings and updates keep the same lifetime as WinTab, independently of the web window.
/// The x86 WPF compatibility window continues to use the same HookManager/SettingsManager APIs.
/// </summary>
internal sealed class DesktopApplication : IApplicationUi, IDisposable
{
    private readonly Dispatcher _dispatcher = Application.Current.Dispatcher;
    private readonly HookManager _hooks;
    private readonly SystemTrayIcon _tray;
    private readonly DesktopUiProcess _ui;
    private readonly DispatcherTimer _autoUpdate;
    private readonly string _version = typeof(DesktopApplication).Assembly.GetName().Version?.ToString(3) ?? "1.0.0";
    private long _revision;
    private bool _disposed;
    private bool _sessionBusy;
    private bool _updateBusy;
    private bool _publishQueued;
    private SessionCommandResult? _sessionResult;
    private Func<string>? _updateFeedback;

    internal DesktopApplication()
    {
        UiStrings.ApplyCulture(SettingsManager.Language);
        _hooks = new HookManager();
        _ui = new DesktopUiProcess(_dispatcher, ExecuteAsync, pid => _hooks.SettingsUiProcessId = pid);
        _tray = new SystemTrayIcon(_hooks, ShowMainWindow, () => Application.Current.Shutdown());
        _tray.StartupChanged += OnStartupChanged;
        _hooks.StateChanged += Publish;
        _hooks.ShellInitialized += Publish;
        _hooks.SessionCommandFinished += OnSessionFinished;
        SettingsManager.StaticPropertyChanged += OnSettingsChanged;
        SettingsManager.StorageErrorChanged += Publish;
        Application.Current.Exit += OnExit;
        _hooks.ApplySettings();
        _autoUpdate = new DispatcherTimer(DispatcherPriority.ApplicationIdle, _dispatcher)
        {
            Interval = TimeSpan.FromSeconds(2)
        };
        _autoUpdate.Tick += CheckAutomatically;
        if (SettingsManager.AutoUpdate) _autoUpdate.Start();
    }

    public void ShowMainWindow() => _ui.Show();
    public Task OpenRecycleBinAsync() => _hooks.OpenRecycleBinAsync();

    private DesktopState State()
    {
        var mismatched = _hooks.MismatchedFeatures;
        return new DesktopState(1, ++_revision, _version, SettingsManager.Snapshot,
            RegistryManager.IsStartupEnabled, _hooks.IsShellReady, SettingsManager.ShouldRecordClosedTabs,
            _sessionBusy, _updateBusy, mismatched.Count == 0 ? null : UiStrings.HookFeaturesMismatched(mismatched),
            _hooks.ShortcutError, _sessionResult is { } result ? UiStrings.SessionResult(result) : null,
            _updateFeedback?.Invoke(), SettingsManager.StorageError is null ? null : UiStrings.DesktopStorageFailed);
    }

    private void Publish()
    {
        if (_disposed) return;
        if (!_dispatcher.CheckAccess())
        {
            _dispatcher.InvokeAsync(Publish);
            return;
        }
        if (_publishQueued) return;
        _publishQueued = true;
        // One setting can raise both preference and hook events. Capture its final state once,
        // after those normal-priority callbacks, instead of repainting for every intermediate notification.
        _dispatcher.InvokeAsync(() =>
        {
            _publishQueued = false;
            if (_disposed) return;
            _tray.RefreshState();
            _ui.Publish(State());
        }, DispatcherPriority.Background);
    }

    private void OnSettingsChanged(object? sender, PropertyChangedEventArgs args)
    {
        if (!_dispatcher.CheckAccess())
        {
            _dispatcher.InvokeAsync(() => OnSettingsChanged(sender, args));
            return;
        }
        var all = string.IsNullOrEmpty(args.PropertyName);
        if (all || args.PropertyName == nameof(SettingsManager.Language))
        {
            UiStrings.ApplyCulture(SettingsManager.Language);
            _tray.ApplyLanguage();
        }
        if (all || args.PropertyName == nameof(SettingsManager.Theme)) ThemeManager.ApplyTheme();
        Publish();
    }

    private void OnStartupChanged(object? sender, EventArgs args) => Publish();
    private void OnSessionFinished(SessionCommandResult result)
    {
        _sessionResult = result;
        _sessionBusy = false;
        Publish();
    }

    private async Task<object> ExecuteAsync(DesktopCommand command)
    {
        if (_disposed) throw new InvalidOperationException("Application is shutting down");
        switch (command.Method)
        {
            case "state": break;
            case "set": Set(command); break;
            case "shortcuts":
                _hooks.ConfigureSessionShortcuts(command.GroupEnabled, command.GroupShortcut,
                    command.TabEnabled, command.TabShortcut);
                break;
            case "restore":
                if (!_sessionBusy)
                {
                    _sessionBusy = true;
                    _sessionResult = null;
                    Publish();
                    try { await _hooks.ExecuteSessionCommandAsync(command.Enabled).ConfigureAwait(true); }
                    finally { _sessionBusy = false; }
                }
                break;
            case "update": await CheckUpdatesAsync().ConfigureAwait(true); break;
            case "logs":
                Directory.CreateDirectory(ExplorerDebugLog.Folder);
                Process.Start(new ProcessStartInfo(ExplorerDebugLog.Folder) { UseShellExecute = true })?.Dispose();
                break;
            case "size": SettingsManager.FormSize = command.Size; break;
            default: throw new InvalidOperationException("Command was not validated");
        }
        return State();
    }

    private void Set(DesktopCommand command)
    {
        var enabled = command.Enabled;
        switch (command.Key)
        {
            case "windowHook": _hooks.SetWindowHook(enabled); break;
            case "reuseTabs": _hooks.SetReuseTabs(enabled); break;
            case "restoreTabs": _hooks.SetRestoreTabs(enabled); break;
            case "restoreSingleTab": _hooks.SetRestoreSingleTab(enabled); break;
            case "restoreOnAnyFolder": _hooks.SetRestoreOnAnyFolder(enabled); break;
            case "reopenClosedTab": _hooks.SetReopenClosedTab(enabled); break;
            case "doubleClickCloseTab": _hooks.SetDoubleClickClose(enabled); break;
            case "doubleClickCloseIncludeNotepad": _hooks.SetDoubleClickCloseIncludeNotepad(enabled); break;
            case "middleClickForegroundTab": _hooks.SetMiddleClickForeground(enabled); break;
            case "wheelSwitchTab": _hooks.SetWheelSwitch(enabled); break;
            case "wheelSwitchSensitivity": _hooks.SetWheelSwitchSensitivity(Enum.Parse<WheelSwitchSensitivity>(command.Text)); break;
            case "showTrayIcon": SettingsManager.ShowTrayIcon = enabled; break;
            case "autoUpdate": SettingsManager.AutoUpdate = enabled; break;
            case "language": SettingsManager.Language = command.Text; break;
            case "theme": SettingsManager.Theme = command.Text; break;
            case "startup":
                if (RegistryManager.IsStartupEnabled != enabled) RegistryManager.ToggleStartup();
                if (RegistryManager.IsStartupEnabled != enabled) throw new IOException("Startup setting was not applied");
                break;
            default: throw new InvalidOperationException("Setting was not validated");
        }
        Publish();
    }

    private async Task CheckUpdatesAsync()
    {
        if (_updateBusy) return;
        _updateBusy = true;
        _updateFeedback = () => UiStrings.UpdateChecking;
        Publish();
        try
        {
            var result = await UpdateManager.CheckForUpdatesWithResultAsync().ConfigureAwait(true);
            _updateFeedback = !result.Completed ? () => UiStrings.UpdateFailed
                : !result.UpdateAvailable ? () => UiStrings.UpdateUpToDate
                : string.IsNullOrWhiteSpace(result.DownloadUrl) ? () => UiStrings.UpdateNoMatchingInstaller
                : () => UiStrings.UpdateOpening;
            if (result.Completed && result.UpdateAvailable && !string.IsNullOrWhiteSpace(result.DownloadUrl))
                UpdateManager.CheckForUpdates();
        }
        finally
        {
            _updateBusy = false;
            Publish();
        }
    }

    private void CheckAutomatically(object? sender, EventArgs args)
    {
        _autoUpdate.Stop();
        if (!_disposed && SettingsManager.AutoUpdate) UpdateManager.CheckForUpdates();
    }

    private void OnExit(object? sender, ExitEventArgs args) => Dispose();

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Application.Current.Exit -= OnExit;
        SettingsManager.StaticPropertyChanged -= OnSettingsChanged;
        SettingsManager.StorageErrorChanged -= Publish;
        _hooks.StateChanged -= Publish;
        _hooks.ShellInitialized -= Publish;
        _hooks.SessionCommandFinished -= OnSessionFinished;
        _tray.StartupChanged -= OnStartupChanged;
        _autoUpdate.Stop();
        _ui.Dispose();
        _tray.Dispose();
        _hooks.Dispose();
        if (!SettingsManager.FlushSettingsAsync(TimeSpan.FromSeconds(1)).GetAwaiter().GetResult())
            Trace.TraceError("Could not confirm settings were saved before exit.");
    }
}
