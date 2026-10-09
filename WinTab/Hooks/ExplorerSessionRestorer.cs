using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace WinTab.Hooks;

/// <param name="RestoredCount">The tabs added to the window, confirmed or not.</param>
internal readonly record struct SessionRestoreResult(int RestoredCount, bool Completed);

/// <summary>
/// Non-destructive restore in saved order. Each new tab must finish navigation before the next request:
/// Explorer publishes its HWND before its folder view exists, and another request can dispatch a native
/// selection into that unfinished view and crash Explorer. Recheck all tabs before finalizing. The initial tab is never
/// navigated, and a normal-launch placeholder is only removed after every addition has been confirmed, while it
/// is still the active tab: Explorer can apply a close command sent to a background tab to its active tab, and
/// closing the active placeholder cannot reach a restored tab either way. The saved active tab is selected last:
/// switching to a tab Explorer is still creating crashes Explorer.
/// </summary>
internal static class ExplorerSessionRestorer
{
    public static async Task<SessionRestoreResult> RestoreAsync(ExplorerSessionRestorePlan plan,
        IExplorerSessionRestoreEnvironment environment, CancellationToken cancellationToken)
    {
        var created = new List<(int SavedIndex, nint Handle, string Location)>();
        foreach (var tab in plan.Tabs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            environment.EnsureUnchanged();
            var handle = await environment.AppendTabAsync(tab.Location);
            if (handle == 0)
                return new SessionRestoreResult(created.Count, false);
            created.Add((tab.SavedIndex, handle, tab.Location));
            cancellationToken.ThrowIfCancellationRequested();
            environment.EnsureUnchanged();
            if (!await environment.ConfirmTabsAsync([(handle, tab.Location)]))
                return new SessionRestoreResult(created.Count, false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        environment.EnsureUnchanged();
        if (created.Count > 0 && !await environment.ConfirmTabsAsync(created.Select(tab => (tab.Handle, tab.Location)).ToArray()))
            return new SessionRestoreResult(created.Count, false);

        cancellationToken.ThrowIfCancellationRequested();
        environment.EnsureUnchanged();
        if (plan.CloseInitialTab && created.Count > 0)
        {
            if (!await environment.CloseInitialTabAsync())
                return new SessionRestoreResult(created.Count, false);
            cancellationToken.ThrowIfCancellationRequested();
            environment.EnsureUnchanged();
        }

        var target = plan.ActivateInitialTab || created.Count == 0
            ? environment.InitialTab
            : created.FirstOrDefault(tab => tab.SavedIndex == plan.ActiveSavedIndex, created[0]).Handle;
        return new SessionRestoreResult(created.Count, await environment.SelectTabAsync(target));
    }
}

internal interface IExplorerSessionRestoreEnvironment
{
    nint InitialTab { get; }
    /// <summary>Abort if the user navigated, changed tabs, changed windows, or retired the operation.</summary>
    void EnsureUnchanged();
    /// <summary>Requests a tab at the location and returns it once Explorer has created it; 0 when it could not be created.</summary>
    Task<nint> AppendTabAsync(string location);
    /// <summary>Waits until every supplied tab has finished loading its location; false when one has not.</summary>
    Task<bool> ConfirmTabsAsync(IReadOnlyList<(nint Tab, string Location)> tabs);
    Task<bool> SelectTabAsync(nint tab);
    /// <summary>Closes the initial tab while it is still active; false when that could not be confirmed.</summary>
    Task<bool> CloseInitialTabAsync();
}
