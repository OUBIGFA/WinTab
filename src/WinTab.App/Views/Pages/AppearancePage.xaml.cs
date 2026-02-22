using System.Windows.Controls;
using WinTab.App.ViewModels;

namespace WinTab.App.Views.Pages;

public partial class AppearancePage : Page
{
    public AppearancePage(AppearanceViewModel viewModel)
    {
        DataContext = viewModel;
        InitializeComponent();
    }
}
