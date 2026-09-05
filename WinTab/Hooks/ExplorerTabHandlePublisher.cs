using System;
using System.Collections.Generic;

namespace WinTab.Hooks;

internal static class ExplorerTabHandlePublisher
{
    public static bool TryGetSingleTabHandle(
        IReadOnlyList<nint> tabHandles,
        Func<nint> getActiveTabHandle,
        out nint tabHandle)
    {
        tabHandle = 0;
        if (tabHandles.Count != 1)
            return false;

        var activeTabHandle = getActiveTabHandle();
        if (activeTabHandle == 0 || activeTabHandle != tabHandles[0])
            return false;

        tabHandle = activeTabHandle;
        return true;
    }
}
