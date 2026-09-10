using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Shell32;

namespace WinTab.Hooks;

/// <summary>
/// Watches the desktop's own "open" action without ever interfering with it. Explorer keeps its
/// native behaviour (open a new window, or simply activate the window whose first tab already shows
/// the folder); this observer only reports which folder the user opened so an already open tab can
/// be brought forward. The default-verb callback returns immediately and never cancels the action.
/// </summary>
internal sealed class ExplorerDesktopOpenObserver : IDisposable
{
    private readonly ShellFolderView _view;
    private readonly Func<bool> _isEnabled;
    private readonly Action<Action> _schedule;
    private readonly Action<string> _folderOpened;
    /// <summary>
    /// Explorer coalesces selection notifications: after rapid changes the next one arrives roughly 250 ms
    /// later, so a notification cannot vouch for a read that starts long after the open. A read that
    /// starts later than this can no longer be tied to the folder the user opened and is skipped.
    /// </summary>
    internal const int OpenReadDeadlineMs = 200;
    private DShellFolderViewEvents_DefaultVerbInvokedEventHandler? _openHandler;
    private readonly DShellFolderViewEvents_SelectionChangedEventHandler _selectionHandler;
    private long _requestVersion;

    public ExplorerDesktopOpenObserver(ShellFolderView view, Func<bool> isEnabled, Action<Action> schedule,
        Action<string> folderOpened)
    {
        _view = view;
        _isEnabled = isEnabled;
        _schedule = schedule;
        _folderOpened = folderOpened;
        _openHandler = OnDefaultVerbInvoked;
        _selectionHandler = OnSelectionChanged;
        _view.DefaultVerbInvoked += _openHandler;
        try
        {
            _view.SelectionChanged += _selectionHandler;
        }
        catch
        {
            _view.DefaultVerbInvoked -= _openHandler;
            throw;
        }
    }

    /// <summary>Returning true lets Explorer continue its own default action; this observer never cancels it.</summary>
    private bool OnDefaultVerbInvoked()
    {
        if (Volatile.Read(ref _openHandler) != null && _isEnabled())
        {
            var requestVersion = Interlocked.Increment(ref _requestVersion);
            var openedAt = Environment.TickCount64;
            try
            {
                _schedule(() => ReportOpenedFolder(requestVersion, openedAt));
            }
            catch (Exception exception) when (exception is ObjectDisposedException or InvalidOperationException or
                TaskSchedulerException { InnerException: ObjectDisposedException or InvalidOperationException })
            {
                ExplorerDebugLog.Write($"Desktop open notification dropped error={exception.GetType().Name}");
            }
        }

        return true;
    }

    // Selection notifications only invalidate pending work; they never read Explorer or delay opening.
    private void OnSelectionChanged() => Interlocked.Increment(ref _requestVersion);

    private bool IsCurrentRequest(long version) => Volatile.Read(ref _openHandler) != null &&
        version == Interlocked.Read(ref _requestVersion) && _isEnabled();

    private void ReportOpenedFolder(long requestVersion, long openedAt)
    {
        if (!IsCurrentRequest(requestVersion))
            return;
        var delay = Environment.TickCount64 - openedAt;
        if (delay > OpenReadDeadlineMs)
        {
            ExplorerDebugLog.Write($"Desktop open selection read skipped; started too late delay={delay}");
            return;
        }

        FolderItems? selection = null;
        FolderItem? item = null;
        try
        {
            selection = _view.SelectedItems();
            if (selection == null || selection.Count != 1)
                return;
            item = selection.Item(0);
            if (item == null || !item.IsFolder || !item.IsFileSystem || item.IsLink)
                return;
            var path = item.Path;
            // A COM read can pump another selection or open notification on this STA thread.
            if (string.IsNullOrWhiteSpace(path) || !IsCurrentRequest(requestVersion))
                return;
            _folderOpened(path);
        }
        finally
        {
            if (item != null && Marshal.IsComObject(item))
                Marshal.ReleaseComObject(item);
            if (selection != null && Marshal.IsComObject(selection))
                Marshal.ReleaseComObject(selection);
        }
    }

    public void Dispose()
    {
        var openHandler = Interlocked.Exchange(ref _openHandler, null);
        if (openHandler == null)
            return;
        try
        {
            DetachEvent(() => _view.DefaultVerbInvoked -= openHandler);
            DetachEvent(() => _view.SelectionChanged -= _selectionHandler);
        }
        finally
        {
            if (Marshal.IsComObject(_view))
                Marshal.ReleaseComObject(_view);
        }
    }

    private static void DetachEvent(Action detach)
    {
        try
        {
            detach();
        }
        catch (Exception exception) when (exception is COMException or InvalidComObjectException)
        {
            ExplorerDebugLog.Write($"Desktop open event connection already closed error={exception.GetType().Name}");
        }
    }
}
