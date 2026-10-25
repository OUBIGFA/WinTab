using SHDocVw;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>File-list location snapshots and parent navigation on the existing Shell STA.</summary>
public partial class ExplorerWatcher
{
    internal static uint FileListBrowseFlags(FileListDoubleClickAction action) => action switch
    {
        FileListDoubleClickAction.Parent => 0x2001, // SBSP_PARENT | SBSP_SAMEBROWSER
        FileListDoubleClickAction.Back => 0x4001, // SBSP_NAVIGATEBACK | SBSP_SAMEBROWSER
        _ => throw new ArgumentOutOfRangeException(nameof(action))
    };

    /// <summary>Lock-free metadata read for the Raw Input thread; this never touches a COM object.</summary>
    internal string? TryGetFileListLocationSnapshot(WindowIdentity window, WindowIdentity tab)
    {
        foreach (var info in Volatile.Read(ref _trackedWindowInfos))
        {
            if (info.Closed || info.Identity != window || info.TabIdentity != tab)
                continue;
            var location = info.Location;
            if (string.IsNullOrWhiteSpace(location) || ExplorerLocationPolicy.IsControlPanel(location))
                return null;
            return window.IsCurrent && tab.IsCurrent && WinApi.GetParent(tab.Handle) == window.Handle &&
                   ExplorerNavigationAccess.ActiveTab(window.Handle) == tab.Handle
                ? Helper.NormalizeLocation(location)
                : null;
        }
        return null;
    }

    /// <summary>
    /// Navigates only the captured active tab. A live Shell location check prevents a native folder double-click
    /// that already changed directory from being followed by an unintended second navigation.
    /// </summary>
    internal async Task<bool> NavigateFileListAsync(ExplorerFileListDoubleClickRequest request,
        FileListDoubleClickAction action, Func<bool> isCurrent)
    {
        var navigated = false;
        try
        {
            await RunShellWorkAsync(async () =>
            {
                if (!isCurrent() || !ExplorerFileListAutomation.IsCurrent(request.Second))
                    return;

                var browser = await FindShellWindowByTabHandle(request.Second.Tab.Handle,
                    request.Second.Window.Handle).ConfigureAwait(true);
                if (browser == null || !isCurrent() || !ExplorerFileListAutomation.IsCurrent(request.Second) ||
                    !AreLocationsEquivalent(TryGetLocation(browser), request.Second.Location))
                    return;

                // ReSharper disable once SuspiciousTypeConversion.Global
                if (browser is not Interop.IServiceProvider serviceProvider)
                    return;
                var queryResult = serviceProvider.QueryService(ref _shellBrowserGuid, ref _shellBrowserGuid,
                    out var shellBrowser);
                if (queryResult < 0 || shellBrowser == null)
                    return;

                try
                {
                    if (shellBrowser.GetWindow(out var actualTab) < 0 || actualTab != request.Second.Tab.Handle ||
                        !isCurrent() || !ExplorerFileListAutomation.IsCurrent(request.Second) ||
                        ExplorerNavigationAccess.ForegroundFrame() != request.Second.Window.Handle)
                        return;

                    // Read again after QueryService: a busy Explorer may have completed a competing navigation.
                    var location = TryGetLocation(browser);
                    if (ExplorerLocationPolicy.IsControlPanel(location) ||
                        !AreLocationsEquivalent(location, request.Second.Location) || !isCurrent())
                        return;

                    var result = shellBrowser.BrowseObject(nint.Zero, FileListBrowseFlags(action));
                    if (result < 0)
                    {
                        ExplorerDebugLog.Write($"File-list {action} navigation rejected hr={result:X8}");
                        return;
                    }
                    navigated = true;
                }
                finally
                {
                    Marshal.ReleaseComObject(shellBrowser);
                }
            }).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            return false;
        }
        catch (Exception exception)
        {
            ExplorerDebugLog.Write($"File-list {action} navigation failed: {exception.GetType().Name}:{exception.Message}");
            return false;
        }
        return navigated;
    }
}
