using System.Windows.Controls;
using WinTab.App.ViewModels;

namespace WinTab.App.Views.Pages;

public partial class GroupsPage : Page
{
    public GroupsPage(GroupsViewModel viewModel)
    {
        DataContext = viewModel;
        InitializeComponent();

        Loaded += (_, _) => viewModel.RefreshCommand.Execute(null);
    }
}
