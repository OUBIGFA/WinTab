using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WinTab.Core.Enums;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Persistence;

namespace WinTab.App.ViewModels;

public partial class AutoGroupingViewModel : ObservableObject
{
    private readonly AppSettings _settings;
    private readonly SettingsStore _settingsStore;
    private readonly Logger _logger;

    /// <summary>
    /// Static array of match types for use in DataGrid combo box binding.
    /// </summary>
    public static AutoGroupMatchType[] MatchTypes { get; } =
        Enum.GetValues<AutoGroupMatchType>();

    public ObservableCollection<AutoGroupRule> Rules { get; }

    [ObservableProperty]
    private AutoGroupRule? _selectedRule;

    [ObservableProperty]
    private bool _autoApply;

    public ObservableCollection<string> Exclusions { get; }

    [ObservableProperty]
    private string? _selectedExclusion;

    [ObservableProperty]
    private string _newExclusionText = string.Empty;

    public AutoGroupingViewModel(
        AppSettings settings,
        SettingsStore settingsStore,
        Logger logger)
    {
        _settings = settings;
        _settingsStore = settingsStore;
        _logger = logger;

        Rules = new ObservableCollection<AutoGroupRule>(settings.AutoGroupRules);
        Exclusions = new ObservableCollection<string>(settings.ExcludedProcesses);
        _autoApply = settings.AutoApplyRules;

        Rules.CollectionChanged += (_, _) => SyncRulesToSettings();
        Exclusions.CollectionChanged += (_, _) => SyncExclusionsToSettings();
    }

    partial void OnAutoApplyChanged(bool value)
    {
        _settings.AutoApplyRules = value;
        SaveSettings();
    }

    [RelayCommand]
    private void AddRule()
    {
        var rule = new AutoGroupRule
        {
            MatchType = AutoGroupMatchType.ProcessName,
            MatchValue = string.Empty,
            GroupName = "New Group",
            Priority = Rules.Count,
            Enabled = true
        };

        Rules.Add(rule);
        SelectedRule = rule;
        _logger.Info("Auto-group rule added.");
    }

    [RelayCommand]
    private void RemoveRule()
    {
        if (SelectedRule is null)
            return;

        Rules.Remove(SelectedRule);
        SelectedRule = null;
        _logger.Info("Auto-group rule removed.");
    }

    [RelayCommand]
    private void AddExclusion()
    {
        string processName = NewExclusionText.Trim();
        if (string.IsNullOrWhiteSpace(processName))
            return;

        if (Exclusions.Contains(processName))
            return;

        Exclusions.Add(processName);
        NewExclusionText = string.Empty;
        _logger.Info($"Exclusion added: {processName}");
    }

    [RelayCommand]
    private void RemoveExclusion()
    {
        if (SelectedExclusion is null)
            return;

        Exclusions.Remove(SelectedExclusion);
        SelectedExclusion = null;
        _logger.Info("Exclusion removed.");
    }

    private void SyncRulesToSettings()
    {
        _settings.AutoGroupRules = [.. Rules];
        SaveSettings();
    }

    private void SyncExclusionsToSettings()
    {
        _settings.ExcludedProcesses = [.. Exclusions];
        SaveSettings();
    }

    private void SaveSettings()
    {
        _settingsStore.SaveDebounced(_settings);
    }
}
