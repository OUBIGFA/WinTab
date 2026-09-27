using Shell32;
using SHDocVw;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Models;

namespace WinTab.Hooks;

/// <summary>
/// Reading, comparing and changing a tab's location and selection.
/// </summary>
public partial class ExplorerWatcher
{
    private const int NavigationCompleteWaitMs = 600;
    private const int NavigationVerificationWaitMs = 1_200;
    private const int StartupLocationCacheLimit = 512;

    private Task<string> ResolveInitialLocation(InternetExplorer window)
    {
        return _locationResolver.ResolveAsync(
            () => TryGetLocation(window),
            IsStartupExplorerLocation,
            isBusy: () => IsShellWindowBusy(window), cancellationToken: CurrentCancellation);
    }
    private static bool IsShellWindowBusy(InternetExplorer window)
    {
        try
        {
            return window.Busy;
        }
        catch (COMException exception) when (!IsDisconnectedShell(exception))
        {
            return false;
        }
    }
    private static string TryGetLocation(InternetExplorer window)
    {
        try
        {
            return GetLocation(window);
        }
        catch (COMException exception) when (!IsDisconnectedShell(exception))
        {
            ExplorerDebugLog.Write($"Tab location temporarily unavailable error={exception.HResult:X8}");
            return string.Empty;
        }
    }
    private static nint SafeGetWindowHandle(InternetExplorer window)
    {
        try
        {
            return new IntPtr(window.HWND);
        }
        catch (COMException exception) when (!IsDisconnectedShell(exception))
        {
            return 0;
        }
    }
    private bool IsStartupExplorerLocation(string location)
    {
        if (string.IsNullOrWhiteSpace(location))
            return true;

        location = Helper.NormalizeLocation(location);
        if (StringComparer.OrdinalIgnoreCase.Equals(location, _defaultLocation))
            return true;

        if (_startupLocationCache.TryGetValue(location, out var cached))
            return cached;

        bool isStartup;
        try
        {
            isStartup = _shellPathComparer.IsEquivalent(location, _defaultLocation);
        }
        catch
        {
            isStartup = false;
        }

        if (_startupLocationCache.Count >= StartupLocationCacheLimit)
            _startupLocationCache.Clear();

        _startupLocationCache[location] = isStartup;
        return isStartup;
    }
    private Task<int> WaitForExplorerTabCount(nint hWnd)
    {
        return Helper.DoUntilConditionAsync(
            () => ExplorerWindowDiscovery.GetAllExplorerTabs(hWnd).Take(2).Count(),
            count => count > 0,
            800,
            20, CurrentCancellation);
    }

    private async Task<bool> WaitForNavigation(InternetExplorer window, string targetLocation, int timeoutMs = 5_000)
    {
        if (string.IsNullOrWhiteSpace(targetLocation))
            return true;

        var resolvedLocation = await Helper.DoUntilConditionAsync(
            () => TryGetLocation(window),
            location => AreLocationsEquivalent(location, targetLocation),
            timeoutMs,
            50, CurrentCancellation);

        return AreLocationsEquivalent(resolvedLocation, targetLocation);
    }
    private async Task<bool> NavigateNewTabToTargetAsync(InternetExplorer window, string targetLocation, bool navigationStarted = false)
    {
        if (string.IsNullOrWhiteSpace(targetLocation))
            return true;

        if (AreLocationsEquivalent(TryGetLocation(window), targetLocation))
            return true;

        // Direct requests must reach their location themselves. Redirecting an unconfirmed tab could
        // overwrite a competing user navigation, so failure is reported without another Navigate2.
        if (navigationStarted)
            return await WaitForNavigation(window, targetLocation, NavigationVerificationWaitMs);

        var navigationCompleted = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        DWebBrowserEvents2_NavigateComplete2EventHandler? navigateHandler = null;
        navigateHandler = (object _, ref object _) =>
        {
            // Trust the first NavigateComplete2 unconditionally. Explorer can fire the event a few
            // ticks before LocationURL is updated, so gating on AreLocationsEquivalent here would
            // leave tcs unset on the typical fast path and stall every merge until WaitForNavigation
            // catches up.
            navigationCompleted.TrySetResult(true);
        };

        try
        {
            window.NavigateComplete2 += navigateHandler;
            if (!await NavigateToTargetIfNeeded(window, targetLocation))
            {
                ExplorerDebugLog.Write($"OpenTab navigate-failed target={targetLocation}");
                return false;
            }

            ExplorerDebugLog.Write($"OpenTab navigated target={targetLocation}");
            await Task.WhenAny(navigationCompleted.Task, Task.Delay(NavigationCompleteWaitMs, CurrentCancellation));
            EnsureCurrentMerge();

            if (AreLocationsEquivalent(TryGetLocation(window), targetLocation))
                return true;

            if (await WaitForNavigation(window, targetLocation, NavigationVerificationWaitMs))
                return true;

            ExplorerDebugLog.Write($"OpenTab target-check-failed target={targetLocation} current={TryGetLocation(window)}");
            return false;
        }
        finally
        {
            if (navigateHandler != null)
                window.NavigateComplete2 -= navigateHandler;
        }
    }
    private async Task<bool> NavigateToTargetIfNeeded(InternetExplorer window, string targetLocation)
    {
        if (string.IsNullOrWhiteSpace(targetLocation))
            return true;

        try
        {
            if (AreLocationsEquivalent(TryGetLocation(window), targetLocation))
                return true;

            await Navigate(window, targetLocation);
            return true;
        }
        catch (COMException exception) when (!IsDisconnectedShell(exception))
        {
            ExplorerDebugLog.Write($"Tab navigation temporarily unavailable error={exception.HResult:X8}");
            return false;
        }
    }
    private bool AreLocationsEquivalent(string location, string targetLocation)
    {
        if (string.IsNullOrWhiteSpace(location) || string.IsNullOrWhiteSpace(targetLocation))
            return false;

        if (StringComparer.OrdinalIgnoreCase.Equals(location, targetLocation))
            return true;

        var normalizedLocation = Helper.NormalizeLocation(location);
        var normalizedTargetLocation = Helper.NormalizeLocation(targetLocation);
        if (StringComparer.OrdinalIgnoreCase.Equals(normalizedLocation, normalizedTargetLocation))
            return true;

        try
        {
            return _shellPathComparer.IsEquivalent(location, targetLocation);
        }
        catch
        {
            return false;
        }
    }
    private bool TryGetRecentlyClosedWindow(string location, out WindowRecord? closedWindow, int maxAge = 2_000)
    {
        if (string.IsNullOrWhiteSpace(location))
        {
            closedWindow = null;
            return false;
        }

        nint targetPidl = 0;
        try
        {
            targetPidl = _shellPathComparer.GetPidlFromPath(location);
            lock (_closedWindowsLock)
            {
                for (var i = _closedWindows.Count - 1; i >= 0; i--)
                {
                    var record = _closedWindows[i];
                    if (Environment.TickCount - record.CreatedAt > maxAge) break;
                    if (!_shellPathComparer.IsEquivalent(location, record.Location, targetPidl)) continue;
                    _closedWindows.RemoveAt(i);
                    closedWindow = record;
                    return true;
                }
            }
            closedWindow = null;
            return false;
        }
        finally
        {
            if (targetPidl != 0)
                Marshal.FreeCoTaskMem(targetPidl);
        }
    }

    private static string[]? GetSelectedItems(InternetExplorer window)
    {
        if (window.Document is not ShellFolderView document)
            return null;
        var selectedItems = document.SelectedItems();
        if (selectedItems == null)
            return null;
        var count = selectedItems.Count;
        if (count == 0) return Array.Empty<string>();

        var result = new string[count];
        for (var index = 0; index < count; index++)
        {
            var item = selectedItems.Item(index);
            if (item == null)
                return null;
            result[index] = item.Name;
        }

        return result;
    }
    private static string[]? TryGetSelectedItems(InternetExplorer window)
    {
        try
        {
            return GetSelectedItems(window);
        }
        catch (COMException exception) when (!IsDisconnectedShell(exception))
        {
            ExplorerDebugLog.Write($"Selection capture unavailable error={exception.GetType().Name}");
            return null;
        }
    }
    private bool SelectItems(InternetExplorer window, string[]? names)
    {
        if (names == null || names.Length == 0) return true;
        var identity = WindowIdentity.Capture(SafeGetWindowHandle(window));
        EnsureWindowIdentity(identity);

        if (window.Document is not ShellFolderView document) return false;

        const int selectItem = 1;
        const int deselectOthers = 4;
        var selectedAny = false;
        foreach (var name in names)
        {
            EnsureWindowIdentity(identity);
            object item = document.Folder.ParseName(name);
            if (item == null) continue;
            document.SelectItem(ref item, selectedAny ? selectItem : selectItem | deselectOthers);
            selectedAny = true;
        }
        return selectedAny;
    }
    private static string GetLocation(InternetExplorer window)
    {
        var path = window.LocationURL;
        if (!string.IsNullOrWhiteSpace(path)) return Helper.NormalizeLocation(path);

        // Recycle Bin, This PC, etc
        if (window.Document is not ShellFolderView document || document.Folder is not Folder2 folder)
            return string.Empty;
        return Helper.NormalizeLocation(folder.Self.Path);
    }
    private async Task Navigate(InternetExplorer window, string path)
    {
        var identity = WindowIdentity.Capture(SafeGetWindowHandle(window));
        EnsureWindowIdentity(identity);
        if (!path.Contains('#') && !path.Contains("%23"))
        {
            window.Navigate2(path);
            return;
        }

        var folder = await RunInStaThread(() =>
        {
            Shell? shell = null;
            Folder? folder;
            try
            {
                shell = new Shell();
                folder = shell.NameSpace(path);
            }
            finally
            {
                if (shell != null)
                    Marshal.ReleaseComObject(shell);
            }
            return folder;
        });

        try
        {
            EnsureWindowIdentity(identity);
            window.Navigate2(folder);
        }
        finally
        {
            if (folder != null)
                Marshal.ReleaseComObject(folder);
        }
    }
}
