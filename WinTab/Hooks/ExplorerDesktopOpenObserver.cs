using System;
using System.Runtime.InteropServices;
using System.Threading;
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
    private DShellFolderViewEvents_DefaultVerbInvokedEventHandler? _openHandler;

    public ExplorerDesktopOpenObserver(ShellFolderView view, Func<bool> isEnabled, Action<Action> schedule,
        Action<string> folderOpened)
    {
        _view = view;
        _isEnabled = isEnabled;
        _schedule = schedule;
        _folderOpened = folderOpened;
        _openHandler = OnDefaultVerbInvoked;
        _view.DefaultVerbInvoked += _openHandler;
    }

    /// <summary>Returning true lets Explorer continue its own default action; this observer never cancels it.</summary>
    private bool OnDefaultVerbInvoked()
    {
        if (Volatile.Read(ref _openHandler) != null && _isEnabled())
        {
            try
            {
                _schedule(ReportOpenedFolder);
            }
            catch (Exception exception) when (exception is ObjectDisposedException or InvalidOperationException)
            {
                ExplorerDebugLog.Write($"Desktop open notification dropped error={exception.GetType().Name}");
            }
        }

        return true;
    }

    private void ReportOpenedFolder()
    {
        if (Volatile.Read(ref _openHandler) == null || !_isEnabled())
            return;

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
            if (string.IsNullOrWhiteSpace(path) || Volatile.Read(ref _openHandler) == null)
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
            _view.DefaultVerbInvoked -= openHandler;
        }
        catch (Exception exception) when (exception is COMException or InvalidComObjectException)
        {
            ExplorerDebugLog.Write($"Desktop open event connection already closed error={exception.GetType().Name}");
        }
        if (Marshal.IsComObject(_view))
            Marshal.ReleaseComObject(_view);
    }
}
