using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;

namespace WinTab.App.ViewModels;

public partial class MainViewModel : ObservableObject
{
    [ObservableProperty]
    private string _currentPageTitle = string.Empty;

    [ObservableProperty]
    private object? _currentPage;

    public ObservableCollection<NavigationItem> NavigationItems { get; } = [];

    public MainViewModel()
    {
        NavigationItems.Add(new NavigationItem("General", "Settings24"));
        NavigationItems.Add(new NavigationItem("Behavior", "Settings24"));
        NavigationItems.Add(new NavigationItem("About", "Info24"));
    }
}

public sealed record NavigationItem(string Title, string IconGlyph);
