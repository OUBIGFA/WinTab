using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WinTab.Core.Interfaces;
using WinTab.Core.Models;
using WinTab.Diagnostics;

namespace WinTab.App.ViewModels;

/// <summary>
/// Lightweight display model for a tab group shown in the groups list.
/// </summary>
public sealed class GroupDisplayItem
{
    public Guid Id { get; init; }
    public string Name { get; init; } = string.Empty;
    public int TabCount { get; init; }
    public string ActiveTabTitle { get; init; } = string.Empty;
}

public partial class GroupsViewModel : ObservableObject
{
    private readonly IGroupManager? _groupManager;
    private readonly Logger _logger;

    public ObservableCollection<GroupDisplayItem> Groups { get; } = [];

    [ObservableProperty]
    private GroupDisplayItem? _selectedGroup;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasNoGroups))]
    private bool _hasGroups;

    public bool HasNoGroups => !HasGroups;

    public GroupsViewModel(Logger logger, IGroupManager? groupManager = null)
    {
        _logger = logger;
        _groupManager = groupManager;
    }

    [RelayCommand]
    private void Refresh()
    {
        Groups.Clear();

        if (_groupManager is null)
        {
            HasGroups = false;
            return;
        }

        try
        {
            var allGroups = _groupManager.GetAllGroups();
            foreach (TabGroup group in allGroups)
            {
                var activeTab = group.Tabs.FirstOrDefault(t => t.Handle == group.ActiveHandle);
                Groups.Add(new GroupDisplayItem
                {
                    Id = group.Id,
                    Name = group.Name,
                    TabCount = group.Tabs.Count,
                    ActiveTabTitle = activeTab?.Title ?? "(none)"
                });
            }

            HasGroups = Groups.Count > 0;
            _logger.Info($"Groups refreshed: {Groups.Count} group(s).");
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to refresh groups.", ex);
            HasGroups = false;
        }
    }

    [RelayCommand]
    private void Disband()
    {
        if (SelectedGroup is null || _groupManager is null)
            return;

        try
        {
            _groupManager.DisbandGroup(SelectedGroup.Id);
            _logger.Info($"Disbanded group: {SelectedGroup.Name}");
            Refresh();
        }
        catch (Exception ex)
        {
            _logger.Error($"Failed to disband group {SelectedGroup.Name}.", ex);
        }
    }

    [RelayCommand]
    private void DisbandAll()
    {
        if (_groupManager is null)
            return;

        try
        {
            var allGroups = _groupManager.GetAllGroups().ToList();
            foreach (TabGroup group in allGroups)
            {
                _groupManager.DisbandGroup(group.Id);
            }

            _logger.Info($"Disbanded all groups ({allGroups.Count}).");
            Refresh();
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to disband all groups.", ex);
        }
    }
}
