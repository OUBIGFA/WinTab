using System;
using System.Collections.Generic;
using System.Linq;

namespace WinTab.Hooks;

/// <summary>One window's MRU order, using tab identities so moves and duplicate titles stay distinct.</summary>
internal sealed class TabActivationHistory
{
    private readonly List<string> _recent = [];

    public void Observe(ExplorerTabAutomation.Tab[] tabs)
    {
        // An incomplete accessibility snapshot must not erase previously observed activations.
        if (tabs.Length == 0 || tabs.Any(tab => string.IsNullOrEmpty(tab.Id)) ||
            tabs.Select(tab => tab.Id).Distinct(StringComparer.Ordinal).Count() != tabs.Length)
            return;
        var live = tabs.Select(tab => tab.Id).ToHashSet(StringComparer.Ordinal);
        _recent.RemoveAll(id => !live.Contains(id));
        var selected = tabs.Where(tab => tab.Selected).ToArray();
        if (selected.Length != 1) return;
        _recent.Remove(selected[0].Id);
        _recent.Insert(0, selected[0].Id);
    }

    /// <summary>Freeze the return order before closing; the native successor is not a user activation.</summary>
    public string[] GetReturnOrder(string closingId) => _recent.Where(id => id != closingId).ToArray();

    public static string? FindReturnTab(string[] order, string closingId, ExplorerTabAutomation.Tab[] before,
        ExplorerTabAutomation.Tab[] after)
    {
        // Do not switch during tab creation, a cancelled save dialog or another concurrent close.
        var expected = before.Where(tab => tab.Id != closingId).Select(tab => tab.Id).ToHashSet(StringComparer.Ordinal);
        if (after.Length != before.Length - 1 || after.Any(tab => string.IsNullOrEmpty(tab.Id)) ||
            !expected.SetEquals(after.Select(tab => tab.Id)))
            return null;
        return order.FirstOrDefault(expected.Contains);
    }
}
