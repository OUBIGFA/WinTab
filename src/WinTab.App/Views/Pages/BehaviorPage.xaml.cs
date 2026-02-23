using System.Windows.Controls;
using WinTab.App.ViewModels;

namespace WinTab.App.Views.Pages;

public partial class BehaviorPage : Page
{
    public BehaviorPage(BehaviorViewModel viewModel)
    {
        DataContext = viewModel;
        InitializeComponent();
    }
}
