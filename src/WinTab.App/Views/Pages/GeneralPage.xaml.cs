using System.Windows.Controls;
using WinTab.App.ViewModels;

namespace WinTab.App.Views.Pages;

public partial class GeneralPage : Page
{
    public GeneralPage(GeneralViewModel viewModel)
    {
        DataContext = viewModel;
        InitializeComponent();
    }
}
