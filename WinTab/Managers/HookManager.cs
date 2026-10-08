using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Hooks;
using WinTab.UI.Localization;

namespace WinTab.Managers;

/// <summary>A setting that is carried out by an input or window hook, which can fail to start or stop.</summary>
public enum HookFeature { MergeWindows, DoubleClickClose, MiddleClickForeground, WheelSwitch }

/// <summary>
/// Single owner of the hook on/off state: every UI surface calls the Set* methods here, which
/// persist the setting, apply it to the hooks, and keep the two coupled toggles consistent.
/// </summary>
public sealed class HookManager : IDisposable
{
    private readonly SynchronizationContext _syncContext;
    private readonly ExplorerWatcher _explorerWatcher;
    private readonly ExplorerTabDoubleClickHook _doubleClickHook;
    private readonly ExplorerNavigationMiddleClickHook _middleClickHook;
    private readonly ExplorerTabWheelSwitchHook _wheelSwitchHook;
    private readonly ExplorerSessionShortcutHook _sessionShortcuts;
    private readonly System.Windows.SessionEndingCancelEventHandler _sessionEndingHandler;
    private bool _disposed;

    public event Action? StateChanged;
    public event Action? ShellInitialized;
    public event Action<SessionCommandResult>? SessionCommandFinished;
    public string? ShortcutError { get; private set; }

    public HookManager()
    {
        _syncContext = SynchronizationContext.Current ?? new SynchronizationContext();

        _explorerWatcher = new ExplorerWatcher(RegistryManager.GetDefaultExplorerLaunchId, RegistryManager.RestoresFolderWindowsAtSignIn);
        _doubleClickHook = new ExplorerTabDoubleClickHook(_explorerWatcher, () => SettingsManager.DoubleClickCloseTab,
            () => SettingsManager.DoubleClickCloseIncludeNotepad);
        _middleClickHook = new ExplorerNavigationMiddleClickHook();
        _wheelSwitchHook = new ExplorerTabWheelSwitchHook(_explorerWatcher, () => SettingsManager.WheelSwitchSensitivity);
        _sessionShortcuts = new ExplorerSessionShortcutHook(action => _ = ExecuteSessionCommandAsync(action == SessionAction.RestoreGroup));
        _sessionShortcuts.Failed += message =>
        {
            ShortcutError = ShortcutUnavailable(message);
            RaiseStateChanged();
        };

        _explorerWatcher.OnShellInitialized += () => _syncContext.Post(_ =>
        {
            if (_disposed) return;
            UpdateRecycleBinRegistration();
            ShellInitialized?.Invoke();
        }, null);
        _doubleClickHook.StatusChanged += message => ReportHookStatus("Double-click close", message);
        _wheelSwitchHook.StatusChanged += message => ReportHookStatus("Wheel switch", message);

        _sessionEndingHandler = (_, _) => Dispose();
        System.Windows.Application.Current.SessionEnding += _sessionEndingHandler;
    }

    public bool IsShellReady => _explorerWatcher.IsShellReady;

    /// <summary>
    /// Features whose hook is not in the state the user chose: turned on but not started, or turned off but
    /// still running. Toggling the setting again retries the change.
    /// </summary>
    public IReadOnlyList<HookFeature> MismatchedFeatures => FindMismatched(
    [
        (HookFeature.MergeWindows, SettingsManager.IsWindowHookActive, _explorerWatcher),
        (HookFeature.DoubleClickClose, SettingsManager.DoubleClickCloseTab, _doubleClickHook),
        (HookFeature.MiddleClickForeground, SettingsManager.MiddleClickForegroundTab, _middleClickHook),
        (HookFeature.WheelSwitch, SettingsManager.WheelSwitchTab, _wheelSwitchHook)
    ]);

    internal static IReadOnlyList<HookFeature> FindMismatched(IEnumerable<(HookFeature Feature, bool Enabled, IHook Hook)> features) =>
        features.Where(feature => feature.Enabled != feature.Hook.IsHookActive).Select(feature => feature.Feature).ToArray();

    public void ApplySettings()
    {
        SetWindowHook(SettingsManager.IsWindowHookActive);
        SetReuseTabs(SettingsManager.ReuseTabs);
        SetRestoreOnAnyFolder(SettingsManager.RestoreOnAnyFolder);
        SetRestoreSingleTab(SettingsManager.RestoreSingleTab);
        SetRestoreTabs(SettingsManager.RestoreTabs);
        _explorerWatcher.EnableSessionCapture();
        SetReopenClosedTab(SettingsManager.ReopenClosedTab);
        SetDoubleClickClose(SettingsManager.DoubleClickCloseTab);
        SetMiddleClickForeground(SettingsManager.MiddleClickForegroundTab);
        SetWheelSwitch(SettingsManager.WheelSwitchTab);
    }

    /// <summary>Turning window merging off also turns tab reuse off; reuse needs the merge hook.</summary>
    public void SetWindowHook(bool enabled)
    {
        SettingsManager.IsWindowHookActive = enabled;
        TryChangeHookStatus(_explorerWatcher, enabled);

        if (!enabled && SettingsManager.ReuseTabs)
        {
            SettingsManager.ReuseTabs = false;
            _explorerWatcher.SetReuseTabs(false);
        }

        UpdateRecycleBinRegistration();
        RaiseStateChanged();
    }

    /// <summary>Turning tab reuse on also turns window merging on; reuse needs the merge hook.</summary>
    public void SetReuseTabs(bool enabled)
    {
        SettingsManager.ReuseTabs = enabled;
        _explorerWatcher.SetReuseTabs(enabled);

        if (enabled && !SettingsManager.IsWindowHookActive)
        {
            SettingsManager.IsWindowHookActive = true;
            TryChangeHookStatus(_explorerWatcher, true);
        }

        UpdateRecycleBinRegistration();
        RaiseStateChanged();
    }

    private void UpdateRecycleBinRegistration() => RecycleBinOpenRegistration.Update(
        !_disposed && SettingsManager.IsWindowHookActive && SettingsManager.ReuseTabs &&
        _explorerWatcher.IsHookActive && _explorerWatcher.IsShellReady);

    /// <summary>A request arriving during startup or shutdown still opens the page through Explorer.</summary>
    public async Task OpenRecycleBinAsync()
    {
        if (!_disposed && _explorerWatcher.IsShellReady && await _explorerWatcher.OpenRecycleBinAsync()) return;
        ExplorerDebugLog.Write("Recycle Bin request using native new-window opening; reuse is not ready");
        OpenRecycleBinNatively();
    }

    internal static void OpenRecycleBinNatively()
    {
        var start = new System.Diagnostics.ProcessStartInfo(
            System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "explorer.exe"))
        { UseShellExecute = false };
        start.ArgumentList.Add("/n,::{645FF040-5081-101B-9F08-00AA002F954E}");
        System.Diagnostics.Process.Start(start)?.Dispose();
    }

    /// <summary>Only controls automatic reopening; capture, buttons and shortcuts remain available when off.</summary>
    public void SetRestoreTabs(bool enabled)
    {
        SettingsManager.RestoreTabs = enabled;
        _explorerWatcher.SetRestoreTabs(enabled);
        RaiseStateChanged();
    }

    public void SetRestoreOnAnyFolder(bool enabled)
    {
        SettingsManager.RestoreOnAnyFolder = enabled;
        _explorerWatcher.SetRestoreOnAnyFolder(enabled);
        RaiseStateChanged();
    }

    /// <summary>Off: single-tab windows neither replace the saved group nor are restored as one.</summary>
    public void SetRestoreSingleTab(bool enabled)
    {
        SettingsManager.RestoreSingleTab = enabled;
        _explorerWatcher.SetRestoreSingleTab(enabled);
        RaiseStateChanged();
    }

    public void SetReopenClosedTab(bool enabled)
    {
        SettingsManager.ReopenClosedTab = enabled;
        ApplySessionShortcuts();
        RaiseStateChanged();
    }

    public bool ConfigureSessionShortcuts(bool groupEnabled, string group, bool tabEnabled, string tab)
    {
        if (!ExplorerShortcut.TryParse(group, out var groupKey) || !ExplorerShortcut.TryParse(tab, out var tabKey))
        {
            ShortcutError = UiStrings.ShortcutInvalid;
            RaiseStateChanged();
            return false;
        }
        if (groupKey == tabKey)
        {
            ShortcutError = UiStrings.ShortcutDuplicate;
            RaiseStateChanged();
            return false;
        }
        SettingsManager.SetSessionShortcuts(groupEnabled, groupKey.ToString(), tabEnabled, tabKey.ToString());
        ApplySessionShortcuts();
        RaiseStateChanged();
        return ShortcutError == null;
    }

    private void ApplySessionShortcuts() => ApplySessionShortcuts(SettingsManager.Snapshot);

    /// <summary>
    /// Apply one settings snapshot so history and bindings agree. An enabled tab shortcut keeps recording
    /// alive without requiring the separate history toggle; automatic-restore options never gate either key.
    /// </summary>
    internal void ApplySessionShortcuts(AppSettings settings)
    {
        _explorerWatcher.SetRecordClosedTabs(settings.ShouldRecordClosedTabs);
        ShortcutError = null;
        var groupValid = ExplorerShortcut.TryParse(settings.RestoreGroupShortcut, out var group);
        var tabValid = ExplorerShortcut.TryParse(settings.ReopenTabShortcut, out var tab);
        if (!groupValid || !tabValid || group == tab)
            ShortcutError = groupValid && tabValid ? UiStrings.ShortcutDuplicate : UiStrings.ShortcutInvalid;
        try
        {
            _sessionShortcuts.Configure(ShortcutError == null && settings.RestoreGroupShortcutEnabled ? group : null,
                ShortcutError == null && settings.ReopenTabShortcutEnabled ? tab : null);
        }
        catch (Exception exception)
        {
            ShortcutError = ShortcutUnavailable(exception.Message);
            ReportHookStatus("Session shortcuts", exception.Message);
        }
    }

    public async Task ExecuteSessionCommandAsync(bool group)
    {
        if (_disposed) return;
        var result = group ? await _explorerWatcher.RestoreLastSessionAsync() : await _explorerWatcher.ReopenClosedTabAsync();
        _syncContext.Post(_ =>
        {
            if (!_disposed) SessionCommandFinished?.Invoke(result);
        }, null);
    }

    public void SetDoubleClickClose(bool enabled)
    {
        SettingsManager.DoubleClickCloseTab = enabled;
        TryChangeHookStatus(_doubleClickHook, enabled);
        RaiseStateChanged();
    }

    /// <summary>The scope is read per event; changing it needs no hook restart.</summary>
    public void SetDoubleClickCloseIncludeNotepad(bool enabled)
    {
        SettingsManager.DoubleClickCloseIncludeNotepad = enabled;
        RaiseStateChanged();
    }

    public void SetMiddleClickForeground(bool enabled)
    {
        SettingsManager.MiddleClickForegroundTab = enabled;
        TryChangeHookStatus(_middleClickHook, enabled);
        RaiseStateChanged();
    }

    public void SetWheelSwitch(bool enabled)
    {
        SettingsManager.WheelSwitchTab = enabled;
        TryChangeHookStatus(_wheelSwitchHook, enabled);
        RaiseStateChanged();
    }

    public void SetWheelSwitchSensitivity(WheelSwitchSensitivity sensitivity)
    {
        SettingsManager.WheelSwitchSensitivity = sensitivity;
        RaiseStateChanged();
    }

    /// <summary>
    /// A feature whose hook cannot be installed stays off until it is toggled again or WinTab restarts, and is
    /// listed in <see cref="MismatchedFeatures"/>. It must not take the other features, the settings window or
    /// WinTab itself down with it.
    /// </summary>
    internal static bool TryChangeHookStatus(IHook hook, bool isActive)
    {
        if (hook.IsHookActive == isActive)
            return true;

        ExplorerDebugLog.Write($"{hook.GetType().Name} {(isActive ? "starting" : "stopping")}");
        try
        {
            if (isActive)
                hook.StartHook();
            else
                hook.StopHook();
            return true;
        }
        catch (Exception exception)
        {
            ExplorerDebugLog.Write($"{hook.GetType().Name} could not be {(isActive ? "started" : "stopped")}: {exception}");
            return false;
        }
    }

    /// <summary>System messages end in a full stop, which the interface's copy never shows.</summary>
    private static string ShortcutUnavailable(string reason) =>
        UiStrings.ShortcutUnavailable + " " + reason.Trim().TrimEnd('.', '。', '…');

    private static void ReportHookStatus(string source, string message) => ExplorerDebugLog.Write($"{source}: {message}");

    private void RaiseStateChanged()
    {
        _syncContext.Post(_ => StateChanged?.Invoke(), null);
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
        UpdateRecycleBinRegistration();
        System.Windows.Application.Current.SessionEnding -= _sessionEndingHandler;
        try { _sessionShortcuts.Dispose(); }
        catch (Exception exception) { ReportHookStatus("Session shortcuts", exception.Message); }
        _wheelSwitchHook.Dispose();
        _middleClickHook.Dispose();
        _doubleClickHook.Dispose();
        _explorerWatcher.Dispose();
        GC.SuppressFinalize(this);
    }
}
