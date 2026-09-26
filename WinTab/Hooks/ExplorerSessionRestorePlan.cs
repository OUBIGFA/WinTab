using System;
using System.Collections.Generic;
using System.Linq;
using WinTab.Helpers;
using WinTab.Models;

namespace WinTab.Hooks;

internal readonly record struct SessionRestoreTab(string Location, int SavedIndex);

/// <summary>
/// A restore never navigates the user's initial tab. In normal-launch mode a plain launch only supplies a
/// placeholder, which stands in for the first saved tab when it already shows that location and is otherwise
/// closed once every saved tab has been added. In any-folder mode every launch keeps the tab it opened, a start
/// page such as This PC included, first and active, and the saved tabs are added after it.
/// </summary>
/// <param name="ActiveSavedIndex">The saved tab to select: the saved active tab, or the first surviving one.</param>
internal sealed record ExplorerSessionRestorePlan(SessionRestoreTab[] Tabs, bool ActivateInitialTab,
    bool CloseInitialTab, int ActiveSavedIndex, int SkippedCount)
{
    /// <param name="restoreSingleTab">Off: a saved window with a single tab is not a group and is not restored.</param>
    public static ExplorerSessionRestorePlan? Create(ExplorerSession session, string initialLocation,
        bool startupLocation, bool onAnyFolder, bool restoreSingleTab, IReadOnlySet<int> availableIndices)
    {
        if (string.IsNullOrWhiteSpace(initialLocation) || !onAnyFolder && !startupLocation ||
            !restoreSingleTab && session.Locations.Length < 2)
            return null;

        // Only the first saved tab can be supplied by a placeholder without changing the saved order.
        var placeholder = startupLocation && !onAnyFolder;
        var initialSavedIndex = placeholder
            ? SameLocation(session.Locations[0], initialLocation) ? 0 : -1
            : Array.FindIndex(session.Locations, location => SameLocation(location, initialLocation));
        var tabs = new List<SessionRestoreTab>();
        var skipped = 0;
        for (var index = 0; index < session.Locations.Length; index++)
        {
            if (index == initialSavedIndex)
                continue;
            if (!availableIndices.Contains(index))
            {
                skipped++;
                continue;
            }
            tabs.Add(new SessionRestoreTab(session.Locations[index], index));
        }

        var present = tabs.Select(tab => tab.SavedIndex).Append(initialSavedIndex).Where(index => index >= 0).ToHashSet();
        var active = present.Contains(session.ActiveTabIndex) ? session.ActiveTabIndex : present.DefaultIfEmpty(-1).Min();
        return new ExplorerSessionRestorePlan(tabs.ToArray(),
            ActivateInitialTab: !placeholder || tabs.Count == 0 || active == initialSavedIndex,
            CloseInitialTab: placeholder && initialSavedIndex < 0 && tabs.Count > 0,
            active, skipped);
    }

    internal static bool SameLocation(string left, string right) =>
        StringComparer.OrdinalIgnoreCase.Equals(Helper.NormalizeLocation(left), Helper.NormalizeLocation(right));
}
