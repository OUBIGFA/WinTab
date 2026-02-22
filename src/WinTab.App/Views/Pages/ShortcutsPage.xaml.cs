using System.Windows.Controls;
using System.Windows.Input;
using WinTab.App.ViewModels;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace WinTab.App.Views.Pages;

public partial class ShortcutsPage : Page
{
    private readonly ShortcutsViewModel _viewModel;

    public ShortcutsPage(ShortcutsViewModel viewModel)
    {
        _viewModel = viewModel;
        DataContext = viewModel;
        InitializeComponent();
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (_viewModel.IsRecording)
        {
            _viewModel.RecordKey(e);
            e.Handled = true;
            return;
        }

        base.OnKeyDown(e);
    }
}
