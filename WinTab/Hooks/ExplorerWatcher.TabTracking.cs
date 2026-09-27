using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using SHDocVw;
using WinTab.Helpers;
using WinTab.Models;
using WinTab.WinAPI;

namespace WinTab.Hooks;

using WindowEntry = DualKeyEntry<InternetExplorer, nint?, WindowInfo>;

public partial class ExplorerWatcher
{
    private volatile bool _registrationRetryPending;

    private static bool IsCurrentTab(WindowInfo info, nint tabHandle) =>
        tabHandle != 0 && info.TabIdentity.Handle == tabHandle && info.TabIdentity.IsCurrent &&
        WinApi.GetParent(tabHandle) == info.Identity.Handle;

    private bool TryGetKnownTabHandle(InternetExplorer window, out nint tabHandle)
    {
        tabHandle = 0;
        if (!TryGetTrackedEntry(window, out var entry) || !IsCurrentWindow(window, entry.Value))
            return false;

        if (entry.OptionalKey is { } cachedHandle and not 0)
        {
            if (!IsCurrentTab(entry.Value, cachedHandle))
                return false;
            tabHandle = cachedHandle;
            return true;
        }

        var parent = entry.Value.Identity.Handle;
        var tabs = ExplorerWindowDiscovery.GetAllExplorerTabs(parent).Take(2).ToArray();
        if (!ExplorerTabHandlePublisher.TryGetSingleTabHandle(tabs, () => GetActiveTabHandle(parent), out var singleTab) ||
            !TryPublishTabHandle(window, singleTab))
            return false;

        tabHandle = singleTab;
        ExplorerDebugLog.Write($"Registered single-tab handle={tabHandle} hwnd={parent}");
        return true;
    }

    private void RemoveExpiredWindowEntries()
    {
        WindowEntry[] entries;
        lock (_windowEntryDictLock)
            entries = ((IEnumerable<WindowEntry>)_windowEntryDict).ToArray();

        foreach (var (window, info, tab) in entries)
        {
            if (!IsCurrentWindow(window, info) || tab.HasValue && !IsCurrentTab(info, tab.Value))
                RemoveWindowAndUnhookEvents(window, info);
        }
    }

    private bool HasPendingTabRegistrations()
    {
        if (_registrationRetryPending)
            return true;

        lock (_windowEntryDictLock)
        {
            foreach (var (window, info, tab) in _windowEntryDict)
            {
                if (!IsCurrentWindow(window, info) || !info.EventsHooked ||
                    !tab.HasValue || !IsCurrentTab(info, tab.Value))
                    return true;
            }
        }
        return false;
    }

    private bool HasUnregisteredExplorerTabs()
    {
        lock (_windowEntryDictLock)
        {
            foreach (var handle in _getExplorerWindows())
            {
                // A frame Explorer has not shown yet (a preloaded frame) is not registered until it is
                // used, so it must not keep full catalog scans running every second.
                if (!WinApi.IsWindowVisible(handle))
                    continue;

                foreach (var tab in ExplorerWindowDiscovery.GetAllExplorerTabs(handle))
                {
                    if (!_windowEntryDict.TryGetValue(tab, out InternetExplorer? window) || window == null ||
                        !_windowEntryDict.TryGetValue(window, out WindowInfo? info) ||
                        info.Identity.Handle != handle || !IsCurrentWindow(window, info) || !IsCurrentTab(info, tab))
                        return true;
                }
            }
        }
        return false;
    }

    private Task<nint> GetTabHandle(InternetExplorer window)
    {
        if (TryGetKnownTabHandle(window, out var handle))
            return Task.FromResult(handle);

        return QueryTabHandle(window, updateDictionary: true);
    }
    private bool TryPublishTabHandle(InternetExplorer window, nint tabHandle)
    {
        if (tabHandle == 0)
            return false;

        try
        {
            lock (_windowEntryDictLock)
            {
                if (!_windowEntryDict.TryGetValue(window, out WindowInfo? info) ||
                    !IsCurrentWindow(window, info) || WinApi.GetParent(tabHandle) != info.Identity.Handle)
                    return false;

                if (_windowEntryDict.TryGetValue(tabHandle, out InternetExplorer? previousWindow) &&
                    previousWindow != null && !ReferenceEquals(previousWindow, window))
                {
                    if (!_windowEntryDict.TryGetValue(previousWindow, out WindowInfo? previousInfo))
                        return false;
                    if (IsCurrentWindow(previousWindow, previousInfo) && IsCurrentTab(previousInfo, tabHandle))
                    {
                        RemoveWindowAndUnhookEvents(window, info, restoreHiddenWindow: false);
                        ExplorerDebugLog.Write($"Registered duplicate tab={tabHandle} hwnd={info.Identity.Handle}");
                        return true;
                    }
                    RemoveWindowAndUnhookEvents(previousWindow, previousInfo);
                }

                var tabIdentity = WindowIdentity.Capture(tabHandle);
                if (!tabIdentity.IsCurrent)
                    return false;
                _windowEntryDict.UpdateOptionalKey(window, tabHandle);
                info.TabIdentity = tabIdentity;
                return true;
            }
        }
        catch (Exception ex) when (!IsDisconnectedShell(ex))
        {
            ExplorerDebugLog.Write($"Tab handle publish failed handle={tabHandle} error={ex.GetType().Name}:{ex.Message}");
            return false;
        }
    }
    private Task<nint> QueryTabHandle(InternetExplorer window, bool updateDictionary)
    {
        return RunInStaThread(() =>
        {
            // ReSharper disable once SuspiciousTypeConversion.Global
            if (window is not Interop.IServiceProvider sp) return 0;

            sp.QueryService(ref _shellBrowserGuid, ref _shellBrowserGuid, out var shellBrowser);
            if (shellBrowser == null) return 0;

            try
            {
                shellBrowser.GetWindow(out var hWnd);
                EnsureCurrentMerge();

                if (updateDictionary && hWnd != 0 && !TryPublishTabHandle(window, hWnd))
                    return 0;

                return hWnd;
            }
            finally
            {
                Marshal.ReleaseComObject(shellBrowser);
            }
        });
    }
    private async Task<InternetExplorer?> FindShellWindowByTabHandle(nint tabHandle, nint parentWindowHandle = 0)
    {
        if (!_staTaskScheduler.IsCurrentThread)
            return await Task.Factory.StartNew(() => FindShellWindowByTabHandle(tabHandle, parentWindowHandle),
                CurrentCancellation, TaskCreationOptions.DenyChildAttach, _staTaskScheduler).Unwrap();

        EnsureCurrentMerge();
        var cachedWindow = GetWindowByTabHandle(tabHandle, parentWindowHandle);
        if (cachedWindow != null)
            return cachedWindow;

        var count = _shellWindows.Count;
        for (var i = count - 1; i >= 0; i--)
        {
            if (_shellWindows.Item(i) is not InternetExplorer window)
                continue;

            if (parentWindowHandle != 0 && SafeGetWindowHandle(window) != parentWindowHandle)
                continue;

            var currentTabHandle = await QueryTabHandle(window, updateDictionary: false);
            EnsureCurrentMerge();
            if (currentTabHandle != tabHandle)
                continue;

            WindowInfo windowInfo;
            InternetExplorer windowToReturn = window;
            lock (_windowEntryDictLock)
            {
                if (_windowEntryDict.TryGetValue(tabHandle, out InternetExplorer? existingWindow) && existingWindow != null &&
                    _windowEntryDict.TryGetValue(existingWindow, out windowInfo!) &&
                    IsCurrentWindow(existingWindow, windowInfo) && IsCurrentTab(windowInfo, tabHandle))
                {
                    windowToReturn = existingWindow;
                }
                else
                {
                    if (!_windowEntryDict.TryGetValue(window, out windowInfo!))
                    {
                        windowInfo = CreateWindowInfo(window);
                        _windowEntryDict.Add(window, windowInfo);
                    }
                    if (!TryPublishTabHandle(window, tabHandle))
                        continue;
                }
            }

            HookWindowEvents(windowToReturn, windowInfo);
            return windowToReturn;
        }

        return null;
    }
    private static nint GetActiveTabHandle(nint windowHandle)
    {
        // Active tab always at the top of the z-index
        return WinApi.FindWindowEx(windowHandle, 0, "ShellTabWindowClass", null);
    }
    private int TrackedWindowCount
    {
        get
        {
            lock (_windowEntryDictLock)
                return _windowEntryDict.Count;
        }
    }
    private bool TryGetTrackedEntry(InternetExplorer window, out WindowEntry entry)
    {
        lock (_windowEntryDictLock)
            return _windowEntryDict.TryGetValue(window, out entry);
    }
    private InternetExplorer? GetWindowByTabHandle(nint tabHandle, nint parentWindowHandle)
    {
        if (tabHandle == 0) return null;

        InternetExplorer? window;
        lock (_windowEntryDictLock)
        {
            if (!_windowEntryDict.TryGetValue(tabHandle, out window) || window == null ||
                !_windowEntryDict.TryGetValue(window, out WindowInfo? info) ||
                !IsCurrentWindow(window, info) || !IsCurrentTab(info, tabHandle))
                return null;

            return parentWindowHandle == 0 || info.Identity.Handle == parentWindowHandle ? window : null;
        }
    }
}
