using System.Windows.Controls;
using WinTab.App.ViewModels;

namespace WinTab.App.Views.Pages;

public partial class UninstallPage : Page
{
    public UninstallPage(UninstallViewModel viewModel)
    {
        DataContext = viewModel;
        InitializeComponent();
    }
}
