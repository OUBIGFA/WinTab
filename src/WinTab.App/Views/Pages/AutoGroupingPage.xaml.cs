using System.Windows.Controls;
using WinTab.App.ViewModels;

namespace WinTab.App.Views.Pages;

public partial class AutoGroupingPage : Page
{
    public AutoGroupingPage(AutoGroupingViewModel viewModel)
    {
        DataContext = viewModel;
        InitializeComponent();
    }
}
