using System;
using System.Threading;
using WinTab.Hooks;

namespace WinTab.Managers;

public sealed class HookManager : IDisposable
{
    private readonly SynchronizationContext _syncContext;
    private readonly ExplorerWatcher _explorerWatcher;
    private readonly ExplorerTabDoubleClickHook _doubleClickHook;

    public event Action? StateChanged;
    public event Action? ShellInitialized;
    public event Action<string>? StatusChanged;

    public HookManager()
    {
        _syncContext = SynchronizationContext.Current ?? new SynchronizationContext();

        _explorerWatcher = new ExplorerWatcher();
        _doubleClickHook = new ExplorerTabDoubleClickHook(_explorerWatcher);

        _explorerWatcher.OnShellInitialized += () => _syncContext.Post(_ => ShellInitialized?.Invoke(), null);
        _explorerWatcher.StatusChanged += message => _syncContext.Post(_ => StatusChanged?.Invoke(message), null);
        _doubleClickHook.StatusChanged += message => _syncContext.Post(_ => StatusChanged?.Invoke(message), null);

        System.Windows.Application.Current.SessionEnding += (_, _) => Dispose();
    }

    public bool IsShellReady => _explorerWatcher.IsShellReady;
    public bool IsWindowHookActive => _explorerWatcher.IsHookActive;
    public bool IsDoubleClickCloseActive => _doubleClickHook.IsHookActive;

    public void ApplySettings()
    {
        SetWindowHook(SettingsManager.IsWindowHookActive);
        SetReuseTabs(SettingsManager.ReuseTabs);
        SetDoubleClickClose(SettingsManager.DoubleClickCloseTab);
    }

    public void SetWindowHook(bool enabled)
    {
        ChangeHookStatus(_explorerWatcher, enabled);

        if (!enabled && SettingsManager.ReuseTabs)
        {
            SettingsManager.ReuseTabs = false;
            _explorerWatcher.SetReuseTabs(false);
        }

        RaiseStateChanged();
    }

    public void SetReuseTabs(bool enabled)
    {
        _explorerWatcher.SetReuseTabs(enabled);

        if (enabled && !SettingsManager.IsWindowHookActive)
        {
            SettingsManager.IsWindowHookActive = true;
            ChangeHookStatus(_explorerWatcher, true);
        }

        RaiseStateChanged();
    }

    public void SetDoubleClickClose(bool enabled)
    {
        ChangeHookStatus(_doubleClickHook, enabled);
        RaiseStateChanged();
    }

    private static void ChangeHookStatus(IHook hook, bool isActive)
    {
        if (hook.IsHookActive == isActive)
            return;

        if (isActive)
            hook.StartHook();
        else
            hook.StopHook();
    }

    private void RaiseStateChanged()
    {
        _syncContext.Post(_ => StateChanged?.Invoke(), null);
    }

    public void Dispose()
    {
        _doubleClickHook.Dispose();
        _explorerWatcher.Dispose();
        GC.SuppressFinalize(this);
    }
}
