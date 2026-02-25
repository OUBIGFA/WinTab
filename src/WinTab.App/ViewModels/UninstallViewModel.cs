using System.Diagnostics;
using System.IO;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.UI.Localization;

namespace WinTab.App.ViewModels;

public sealed partial class UninstallViewModel : ObservableObject
{
    private readonly Logger _logger;
    private readonly string _appDirectory;
    private readonly string? _uninstallerPath;
    private readonly bool _isPortable;

    private const string RemoveUserDataArgument = "/REMOVEUSERDATA=1";

    [ObservableProperty]
    private string _modeText;

    [ObservableProperty]
    private string _uninstallerPathText;

    [ObservableProperty]
    private bool _removeUserDataOnUninstall;

    public bool IsRemoveUserDataOptionEnabled => !_isPortable;

    public UninstallViewModel(Logger logger)
    {
        _logger = logger;

        _appDirectory = Path.GetDirectoryName(Environment.ProcessPath)
            ?? AppContext.BaseDirectory;

        _uninstallerPath = ResolveUninstallerPath(_appDirectory);

        _isPortable = AppPaths.IsPortable;
        _modeText = LocalizationManager.GetString(
            _isPortable ? "Uninstall_ModePortable" : "Uninstall_ModeInstalled");

        _uninstallerPathText = _uninstallerPath
            ?? LocalizationManager.GetString("Uninstall_PathMissing");
    }

    [RelayCommand]
    private void StartUninstall()
    {
        if (_isPortable)
        {
            string title = LocalizationManager.GetString("Uninstall_ConfirmTitle");
            string message = LocalizationManager.GetString("Uninstall_PortableHint");
            System.Windows.MessageBox.Show(
                message,
                title,
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
            OpenAppDirectory();
            return;
        }

        if (string.IsNullOrWhiteSpace(_uninstallerPath) || !File.Exists(_uninstallerPath))
        {
            string title = LocalizationManager.GetString("Uninstall_ConfirmTitle");
            string message = LocalizationManager.GetString("Uninstall_LaunchFailed");
            System.Windows.MessageBox.Show(
                message,
                title,
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning);
            OpenSystemUninstall();
            return;
        }

        string confirmTitle = LocalizationManager.GetString("Uninstall_ConfirmTitle");
        string confirmMessage = LocalizationManager.GetString("Uninstall_Confirm");

        System.Windows.MessageBoxResult result = System.Windows.MessageBox.Show(
            confirmMessage,
            confirmTitle,
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning,
            System.Windows.MessageBoxResult.No);

        if (result != System.Windows.MessageBoxResult.Yes)
            return;

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = _uninstallerPath,
                Arguments = RemoveUserDataOnUninstall ? RemoveUserDataArgument : string.Empty,
                WorkingDirectory = _appDirectory,
                UseShellExecute = true,
            });

            _logger.Info($"Started uninstaller: {_uninstallerPath}");
            App.RequestExplicitShutdown();
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to start uninstaller.", ex);

            string title = LocalizationManager.GetString("Uninstall_ConfirmTitle");
            string message = LocalizationManager.GetString("Uninstall_LaunchFailed");
            System.Windows.MessageBox.Show(
                message,
                title,
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void OpenSystemUninstall()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "ms-settings:appsfeatures",
                UseShellExecute = true,
            });

            return;
        }
        catch
        {
            // Fallback below.
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "control.exe",
                Arguments = "appwiz.cpl",
                UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to open system uninstall entry points.", ex);
        }
    }

    [RelayCommand]
    private void OpenAppDirectory()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = _appDirectory,
                UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to open app directory.", ex);
        }
    }

    private static string? ResolveUninstallerPath(string appDirectory)
    {
        string preferred = Path.Combine(appDirectory, "unins000.exe");
        if (File.Exists(preferred))
            return preferred;

        try
        {
            return Directory
                .GetFiles(appDirectory, "unins*.exe", SearchOption.TopDirectoryOnly)
                .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
        }
        catch
        {
            return null;
        }
    }
}
