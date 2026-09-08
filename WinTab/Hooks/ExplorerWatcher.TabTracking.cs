using System.Collections.Generic;
using System.Linq;
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
}
