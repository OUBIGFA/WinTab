using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Shell32;
using SHDocVw;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

public partial class ExplorerWatcher
{
    private ExplorerDesktopOpenObserver? _desktopOpenObserver;

    private bool CanReuseDesktopFolder => _isForcingTabs && _reuseTabs && !_disposed &&
        _preExistingExplorerWindowsProtected && !_shellLifetime.IsCancellationRequested && !Helper.IsCtrlShiftDown();

    /// <summary>
    /// When a folder is already the first tab of a window, Explorer does not open anything new: it only
    /// activates that window, so the normal window-registration flow never runs. Observing the desktop's
    /// open action lets WinTab bring the matching tab forward while Explorer keeps its native behaviour.
    /// </summary>
    private void ObserveDesktopFolderOpen()
    {
        object location = 0;
        object root = 0;
        InternetExplorer? desktop = null;
        ShellFolderView? view = null;
        try
        {
            desktop = _shellWindows.FindWindowSW(ref location, ref root, 8, out _, 1) as InternetExplorer;
            if (desktop?.Document is not ShellFolderView desktopView)
            {
                ReportStatus("Desktop folder open notifications are unavailable; window merging remains active.");
                return;
            }
            view = desktopView;
            _desktopOpenObserver = new ExplorerDesktopOpenObserver(view, () => CanReuseDesktopFolder,
                ScheduleDesktopWork, location => _ = TryReuseDesktopFolderAsync(location));
            view = null;
        }
        catch (Exception exception) when (exception is COMException or InvalidComObjectException)
        {
            ReportStatus($"Desktop folder open connection failed ({exception.HResult:X8}); window merging remains active.");
        }
        finally
        {
            if (view != null)
                Marshal.ReleaseComObject(view);
            if (desktop != null)
                Marshal.ReleaseComObject(desktop);
        }
    }

    private void ReleaseDesktopFolderOpenObserver()
    {
        var observer = _desktopOpenObserver;
        _desktopOpenObserver = null;
        observer?.Dispose();
    }

    private void ScheduleDesktopWork(Action action)
    {
        _ = RunShellWorkAsync(() =>
        {
            try
            {
                action();
            }
            catch (COMException exception) when (!IsDisconnectedShell(exception))
            {
                ExplorerDebugLog.Write($"Desktop selection temporarily unavailable error={exception.HResult:X8}");
            }
            return Task.CompletedTask;
        });
    }

    /// <summary>Brings an already open tab forward for a folder the user just opened from the desktop.</summary>
    private Task TryReuseDesktopFolderAsync(string location) => RunShellWorkAsync(async () =>
    {
        if (!CanReuseDesktopFolder || TrackedWindowCount == 0)
            return;
        if (!TrySearchForTab(location, 0, out var tab, out var window) || window == null ||
            !TryGetTrackedEntry(window, out var entry))
        {
            ExplorerDebugLog.Write($"Desktop open not-open target={location}");
            return;
        }

        var info = entry.Value;
        var parent = info.Identity.Handle;
        if (WinApi.GetParent(tab) != parent)
            return;
        if (GetActiveTabHandle(parent) == tab)
        {
            ExplorerDebugLog.Write($"Desktop open already-active hwnd={parent} tab={tab} target={location}");
            return;
        }

        var hookGeneration = _hookGeneration;
        bool IsOwnerCurrent() => CanReuseDesktopFolder && hookGeneration == _hookGeneration && IsCurrentWindow(window, info);
        using var operation = new MergeOperation(info.Identity, hookGeneration, _shellLifetime.Token,
            IsOwnerCurrent, MergeTimeoutMs, _hookLifetime.Token);
        await _toOpenWindowsLock.WaitAsync(operation.Token);
        var previousOperation = _currentMerge.Value;
        _currentMerge.Value = operation;
        try
        {
            operation.ThrowIfInvalid();
            var selected = await SelectTabByHandle(parent, tab);
            ExplorerDebugLog.Write($"Desktop open selected={selected} hwnd={parent} tab={tab} target={location}");
        }
        finally
        {
            _currentMerge.Value = previousOperation;
            _toOpenWindowsLock.Release();
        }
    });
}
