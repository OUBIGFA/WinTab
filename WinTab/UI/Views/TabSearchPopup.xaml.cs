using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using System.Collections.Generic;
using WinTab.Helpers;
using WinTab.Models;
using WinTab.Hooks;
using WinTab.Managers;
using Keyboard = System.Windows.Input.Keyboard;

namespace WinTab.UI.Views;

// ReSharper disable once RedundantExtendsListEntry
public partial class TabSearchPopup : Window
{
    private bool _isShowingDialog;
    private bool _isClosing;
    private readonly ExplorerWatcher? _explorerWatcher;
    private readonly TabEngine? _tabEngine;
    private IReadOnlyCollection<WindowRecord> _allWindows = Array.Empty<WindowRecord>();
    private IReadOnlyCollection<WindowRecord> _filteredWindows = Array.Empty<WindowRecord>();
    private IReadOnlyCollection<WindowSearchItem> _allItems = Array.Empty<WindowSearchItem>();
    private IReadOnlyCollection<WindowSearchItem> _filteredItems = Array.Empty<WindowSearchItem>();

    public TabSearchPopup(ExplorerWatcher explorerWatcher)
    {
        InitializeComponent();

        _explorerWatcher = explorerWatcher;
        _tabEngine = null;

        LoadWindows();
        SearchBox.Focus();
    }

    public TabSearchPopup(TabEngine tabEngine)
    {
        InitializeComponent();

        _tabEngine = tabEngine;
        _explorerWatcher = null;

        LoadWindows();
        SearchBox.Focus();
    }

    private void LoadWindows()
    {
        if (_tabEngine != null)
        {
            _allItems = _tabEngine.GetAllWindows()
                .Select(w => new WindowSearchItem
                {
                    Handle = w.Handle,
                    Name = w.Title,
                    DisplayLocation = w.HostType == WindowHostType.Explorer ? w.Title : w.ProcessPath,
                    Location = w.ProcessPath,
                    HostType = w.HostType
                })
                .ToList();

            _filteredItems = _allItems;
            TabsList.ItemsSource = _filteredItems;

            if (_filteredItems.Count > 0)
                TabsList.SelectedIndex = 0;
            return;
        }

        if (_explorerWatcher == null)
            return;

        _allWindows = _explorerWatcher.GetWindows();
        _filteredWindows = _allWindows;
        TabsList.ItemsSource = _filteredWindows;

        if (_filteredWindows.Count > 0)
            TabsList.SelectedIndex = 0;
    }

    private void FilterWindows(string searchText)
    {
        if (_tabEngine != null)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                _filteredItems = _allItems;
            }
            else
            {
                const StringComparison sc = StringComparison.OrdinalIgnoreCase;
                _filteredItems = _allItems
                    .Where(w => w.Name.IndexOf(searchText, sc) != -1 || w.DisplayLocation.IndexOf(searchText, sc) != -1)
                    .OrderByDescending(w => w.Name.IndexOf(searchText, sc) != -1)
                    .ToList();
            }

            TabsList.ItemsSource = _filteredItems;
            if (_filteredItems.Count > 0)
                TabsList.SelectedIndex = 0;
            return;
        }

        if (string.IsNullOrWhiteSpace(searchText))
        {
            _filteredWindows = _allWindows;
        }
        else
        {
            const StringComparison sc = StringComparison.OrdinalIgnoreCase;
            _filteredWindows = _allWindows
                .Where(w => w.Name.IndexOf(searchText, sc) != -1 || w.Location.IndexOf(searchText, sc) != -1)
                .OrderByDescending(w => w.Name.IndexOf(searchText, sc) != -1) // Name matches first
                .ToList();
        }

        TabsList.ItemsSource = _filteredWindows;

        if (_filteredWindows.Count > 0)
            TabsList.SelectedIndex = 0;
    }

    private void SelectNext()
    {
        if (_tabEngine != null)
        {
            if (_filteredItems.Count == 0)
                return;

            if (TabsList.SelectedIndex < _filteredItems.Count - 1)
                TabsList.SelectedIndex++;
            else
                TabsList.SelectedIndex = 0;

            TabsList.ScrollIntoView(TabsList.SelectedItem);
            return;
        }

        if (_filteredWindows.Count == 0)
            return;

        if (TabsList.SelectedIndex < _filteredWindows.Count - 1)
            TabsList.SelectedIndex++;
        else
            TabsList.SelectedIndex = 0; // Wrap around to the first item

        TabsList.ScrollIntoView(TabsList.SelectedItem);
    }

    private void SelectPrevious()
    {
        if (_tabEngine != null)
        {
            if (_filteredItems.Count == 0)
                return;

            if (TabsList.SelectedIndex > 0)
                TabsList.SelectedIndex--;
            else
                TabsList.SelectedIndex = _filteredItems.Count - 1;

            TabsList.ScrollIntoView(TabsList.SelectedItem);
            return;
        }

        if (_filteredWindows.Count == 0)
            return;

        if (TabsList.SelectedIndex > 0)
            TabsList.SelectedIndex--;
        else
            TabsList.SelectedIndex = _filteredWindows.Count - 1; // Wrap around to the last item

        TabsList.ScrollIntoView(TabsList.SelectedItem);
    }

    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        FilterWindows(SearchBox.Text);
    }

    private void SearchBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Down)
        {
            SelectNext();
            e.Handled = true;
        }
        else if (e.Key == Key.Up)
        {
            SelectPrevious();
            e.Handled = true;
        }
        else if (e.Key == Key.Tab)
        {
            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
                SelectPrevious();
            else
                SelectNext();

            e.Handled = true;
        }
        else if (e.Key == Key.Enter && TabsList.SelectedItem != null)
        {
            SwitchToSelectedTab();
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            CloseWindow();
            e.Handled = true;
        }
    }

    private void Window_Deactivated(object sender, EventArgs e)
    {
        if (!_isShowingDialog)
            CloseWindow();
    }

    private void TabItem_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (TabsList.SelectedItem != null)
            SwitchToSelectedTab();
    }

    private void DragHandle_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        DragMove();
    }

    private void SwitchToSelectedTab()
    {
        if (_tabEngine != null)
        {
            if (TabsList.SelectedItem is not WindowSearchItem selected)
                return;

            var entry = _tabEngine.GetAllWindows().FirstOrDefault(w => w.Handle == selected.Handle);
            if (entry != null)
                _tabEngine.Activate(entry);

            CloseWindow();
            return;
        }

        if (TabsList.SelectedItem is not WindowRecord selectedWindow)
            return;

        var asTab = true;
        var duplicate = false;

        // Check if SHIFT is pressed (open as new window)
        if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift))
            asTab = false;

        // Check if CTRL is pressed (duplicate)
        if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            duplicate = true;

        if (_explorerWatcher == null)
            return;

        _ = _explorerWatcher.SwitchTo(
            selectedWindow.Location,
            selectedWindow.Handle,
            selectedWindow.SelectedItems,
            asTab,
            duplicate
        );

        CloseWindow();
    }

    private void CloseWindow()
    {
        if (_isClosing) return;
        _isClosing = true;
        Close();
    }

    private void ClearClosedWindows_Click(object sender, RoutedEventArgs e)
    {
        if (_tabEngine != null)
        {
            // No closed-window history for general windows yet
            return;
        }

        if (_explorerWatcher == null)
            return;

        _isShowingDialog = true;
        var result = CustomMessageBox.Show(
            (string)Application.Current.FindResource("Message_ClearClosedHistory"),
            (string)Application.Current.FindResource("Message_ConfirmClearHistory"),
            MessageBoxButton.YesNo,
            MessageBoxImage.Question,
            MessageBoxResult.No);
        _isShowingDialog = false;

        if (result == MessageBoxResult.Yes)
        {
            _explorerWatcher.ClearClosedWindows();
            LoadWindows();
        }

        // Re-activate the window and focus the search box
        Activate();
        SearchBox.Focus();
    }

    public new void Show()
    {
        base.Show();
        if (Activate()) return;

        Helper.BypassWinForegroundRestrictions();
        Activate();
    }
}

