using System.Windows.Controls;
using WinTab.App.ViewModels;

namespace WinTab.App.Views.Pages;

public partial class AboutPage : Page
{
    public AboutPage(AboutViewModel viewModel)
    {
        DataContext = viewModel;
        InitializeComponent();
    }
}
