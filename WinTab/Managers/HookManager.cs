using System;
using System.Threading;
using WinTab.Hooks;

namespace WinTab.Managers;

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
    private readonly System.Windows.SessionEndingCancelEventHandler _sessionEndingHandler;
    private bool _disposed;

    public event Action? StateChanged;
    public event Action? ShellInitialized;
    public event Action<string>? StatusChanged;

    public HookManager()
    {
        _syncContext = SynchronizationContext.Current ?? new SynchronizationContext();

        _explorerWatcher = new ExplorerWatcher(RegistryManager.GetDefaultExplorerLaunchId);
        _doubleClickHook = new ExplorerTabDoubleClickHook(_explorerWatcher, () => SettingsManager.DoubleClickCloseTab);
        _middleClickHook = new ExplorerNavigationMiddleClickHook();
        _wheelSwitchHook = new ExplorerTabWheelSwitchHook(_explorerWatcher, () => SettingsManager.WheelSwitchSensitivity);

        _explorerWatcher.OnShellInitialized += () => _syncContext.Post(_ => ShellInitialized?.Invoke(), null);
        _explorerWatcher.StatusChanged += message => _syncContext.Post(_ => StatusChanged?.Invoke(message), null);
        _doubleClickHook.StatusChanged += message => ReportHookStatus("Double-click close", message);
        _wheelSwitchHook.StatusChanged += message => ReportHookStatus("Wheel switch", message);
        _middleClickHook.StatusChanged += message => _syncContext.Post(_ => StatusChanged?.Invoke(message), null);

        _sessionEndingHandler = (_, _) => Dispose();
        System.Windows.Application.Current.SessionEnding += _sessionEndingHandler;
    }

    public bool IsShellReady => _explorerWatcher.IsShellReady;
    public bool IsWindowHookActive => _explorerWatcher.IsHookActive;
    public bool IsDoubleClickCloseActive => _doubleClickHook.IsHookActive;

    public void ApplySettings()
    {
        SetWindowHook(SettingsManager.IsWindowHookActive);
        SetReuseTabs(SettingsManager.ReuseTabs);
        SetRestoreOnAnyFolder(SettingsManager.RestoreOnAnyFolder);
        SetRestoreSingleTab(SettingsManager.RestoreSingleTab);
        SetRestoreTabs(SettingsManager.RestoreTabs);
        SetDoubleClickClose(SettingsManager.DoubleClickCloseTab);
        SetMiddleClickForeground(SettingsManager.MiddleClickForegroundTab);
        SetWheelSwitch(SettingsManager.WheelSwitchTab);
    }

    /// <summary>Turning window merging off also turns tab reuse off; reuse needs the merge hook.</summary>
    public void SetWindowHook(bool enabled)
    {
        SettingsManager.IsWindowHookActive = enabled;
        ChangeHookStatus(_explorerWatcher, enabled);

        if (!enabled && SettingsManager.ReuseTabs)
        {
            SettingsManager.ReuseTabs = false;
            _explorerWatcher.SetReuseTabs(false);
        }

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
            ChangeHookStatus(_explorerWatcher, true);
        }

        RaiseStateChanged();
    }

    /// <summary>Restoration and its capture lifecycle remain independent of the merge/reuse toggle pair.</summary>
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

    public void SetDoubleClickClose(bool enabled)
    {
        SettingsManager.DoubleClickCloseTab = enabled;
        ChangeHookStatus(_doubleClickHook, enabled);
        RaiseStateChanged();
    }

    public void SetMiddleClickForeground(bool enabled)
    {
        SettingsManager.MiddleClickForegroundTab = enabled;
        ChangeHookStatus(_middleClickHook, enabled);
        RaiseStateChanged();
    }

    public void SetWheelSwitch(bool enabled)
    {
        SettingsManager.WheelSwitchTab = enabled;
        ChangeHookStatus(_wheelSwitchHook, enabled);
        RaiseStateChanged();
    }

    public void SetWheelSwitchSensitivity(WheelSwitchSensitivity sensitivity)
    {
        SettingsManager.WheelSwitchSensitivity = sensitivity;
        RaiseStateChanged();
    }

    private static void ChangeHookStatus(IHook hook, bool isActive)
    {
        if (hook.IsHookActive == isActive)
            return;

        ExplorerDebugLog.Write($"{hook.GetType().Name} {(isActive ? "starting" : "stopping")}");
        try
        {
            if (isActive)
                hook.StartHook();
            else
                hook.StopHook();
        }
        catch (Exception exception)
        {
            ExplorerDebugLog.Write($"{hook.GetType().Name} could not be {(isActive ? "started" : "stopped")}: {exception}");
            throw;
        }
    }

    private void ReportHookStatus(string source, string message)
    {
        ExplorerDebugLog.Write($"{source}: {message}");
        _syncContext.Post(_ => StatusChanged?.Invoke(message), null);
    }

    private void RaiseStateChanged()
    {
        _syncContext.Post(_ => StateChanged?.Invoke(), null);
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
        System.Windows.Application.Current.SessionEnding -= _sessionEndingHandler;
        _wheelSwitchHook.Dispose();
        _middleClickHook.Dispose();
        _doubleClickHook.Dispose();
        _explorerWatcher.Dispose();
        GC.SuppressFinalize(this);
    }
}
