using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WinTab.Models;

internal static class SelectionSnapshotTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("new windows capture selection before the first periodic scan", CapturesInitialSelection);
        yield return ("an unavailable selection read preserves the last known selection", UnavailableReadPreservesSelection);
        yield return ("an explicitly empty selection clears the old snapshot", EmptySelectionClearsSnapshot);
        yield return ("quick-close records retain the most recently captured selection", QuickCloseUsesLatestSelection);
        yield return ("selection capture skips windows that already closed", ClosedWindowIsNotQueried);
        yield return ("selection capture discards results after a location change", NavigationDiscardsSelection);
        yield return ("selection capture discards results after the window closes", ClosureDiscardsSelection);
    }

    private static Task CapturesInitialSelection()
    {
        var info = new WindowInfo { Location = "C:/Folder" };
        info.RefreshSelection(() => ["initial.txt"], () => true);
        Check.That(info.SelectedItems?.SequenceEqual(["initial.txt"]) == true,
            "A newly registered window must not depend on a previous timer tick.");
        return Task.CompletedTask;
    }

    private static Task UnavailableReadPreservesSelection()
    {
        var info = new WindowInfo { Location = "C:/Folder", SelectedItems = ["known.txt"] };
        info.RefreshSelection(() => null, () => true);
        Check.That(info.SelectedItems?.SequenceEqual(["known.txt"]) == true,
            "A disconnected folder view must not erase the last valid selection.");
        return Task.CompletedTask;
    }

    private static Task EmptySelectionClearsSnapshot()
    {
        var info = new WindowInfo { Location = "C:/Folder", SelectedItems = ["old.txt"] };
        info.RefreshSelection(() => [], () => true);
        Check.That(info.SelectedItems is { Length: 0 }, "Deselecting everything must not resurrect old items.");
        return Task.CompletedTask;
    }

    private static Task QuickCloseUsesLatestSelection()
    {
        var info = new WindowInfo { Location = "C:/Folder", SelectedItems = ["old.txt"] };
        info.RefreshSelection(() => ["latest.txt"], () => true);
        var record = new WindowRecord(info.Location, selectedItems: info.SelectedItems);
        Check.That(record.SelectedItems?.SequenceEqual(["latest.txt"]) == true,
            "A quick reopen must use the last capture, not the older timer snapshot.");
        return Task.CompletedTask;
    }

    private static Task ClosedWindowIsNotQueried()
    {
        var info = new WindowInfo { Closed = true, SelectedItems = ["known.txt"] };
        var queried = false;
        info.RefreshSelection(() =>
        {
            queried = true;
            return ["stale.txt"];
        }, () => !info.Closed);
        Check.That(!queried, "A closed or replaced window must not receive another selection query.");
        return Task.CompletedTask;
    }

    private static Task NavigationDiscardsSelection()
    {
        var info = new WindowInfo { Location = "C:/Folder" };
        info.RefreshSelection(() =>
        {
            info.Location = "C:/Other";
            return ["wrong-folder.txt"];
        }, () => true);
        Check.That(info.SelectedItems == null, "A capture must not attach old-folder items to a new location.");
        return Task.CompletedTask;
    }

    private static Task ClosureDiscardsSelection()
    {
        var info = new WindowInfo { Location = "C:/Folder", SelectedItems = ["known.txt"] };
        info.RefreshSelection(() =>
        {
            info.Closed = true;
            return ["stale.txt"];
        }, () => !info.Closed);
        Check.That(info.SelectedItems?.SequenceEqual(["known.txt"]) == true,
            "Closing during a query must keep the earlier valid snapshot.");
        return Task.CompletedTask;
    }
}
