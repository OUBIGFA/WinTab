using System.Diagnostics;
using System.IO;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WinTab.App.Services;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.Platform.Win32;
using WinTab.UI.Localization;

namespace WinTab.App.ViewModels;

public sealed partial class UninstallViewModel : ObservableObject
{
    private readonly Logger _logger;
    private readonly string _appDirectory;
    private readonly string? _uninstallerPath;
    private readonly bool _isPortable;
    private readonly RegistryOpenVerbInterceptor _openVerbInterceptor;
    private readonly StartupRegistrar _startupRegistrar;

    private const string RemoveUserDataArgument = "/REMOVEUSERDATA=1";

    [ObservableProperty]
    private string _modeText;

    [ObservableProperty]
    private string _uninstallerPathText;

    [ObservableProperty]
    private bool _removeUserDataOnUninstall;

    public bool IsPortable => _isPortable;
    public bool IsInstalled => !_isPortable;
    public bool IsRemoveUserDataOptionEnabled => !_isPortable;

    public UninstallViewModel(Logger logger, RegistryOpenVerbInterceptor openVerbInterceptor, StartupRegistrar startupRegistrar)
    {
        _logger = logger;
        _openVerbInterceptor = openVerbInterceptor;
        _startupRegistrar = startupRegistrar;

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
    private void RestoreSystemConfig()
    {
        string confirmTitle = LocalizationManager.GetString("Uninstall_ConfirmTitle");
        string confirmMessage = LocalizationManager.GetString("Uninstall_RestoreSystemConfig_Confirm");

        MessageBoxResult result = System.Windows.MessageBox.Show(
            confirmMessage,
            confirmTitle,
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);

        if (result != MessageBoxResult.Yes)
            return;

        int failureCount = 0;

        try
        {
            _startupRegistrar.SetEnabled(false);
        }
        catch (Exception ex)
        {
            failureCount++;
            _logger.Error("Failed to remove startup entry during portable cleanup.", ex);
        }

        try
        {
            _openVerbInterceptor.DisableAndRestore();
        }
        catch (Exception ex)
        {
            failureCount++;
            _logger.Error("Failed to restore Explorer open-verb state during portable cleanup.", ex);
            UninstallCleanupHandler.TryRestoreExplorerOpenVerbDefaults(_logger);
        }

        // Clean up the WinTab registry tree (backup cache and any other entries under HKCU\Software\WinTab).
        UninstallCleanupHandler.TryDeleteWinTabRegistryTree();

        string resultKey = failureCount == 0
            ? "Uninstall_RestoreSystemConfig_Success"
            : "Uninstall_RestoreSystemConfig_Partial";

        System.Windows.MessageBox.Show(
            LocalizationManager.GetString(resultKey),
            LocalizationManager.GetString("Uninstall_ConfirmTitle"),
            MessageBoxButton.OK,
            failureCount == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
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
                MessageBoxButton.OK,
                MessageBoxImage.Information);
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
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            OpenSystemUninstall();
            return;
        }

        string confirmTitle = LocalizationManager.GetString("Uninstall_ConfirmTitle");
        string confirmMessage = LocalizationManager.GetString("Uninstall_Confirm");

        MessageBoxResult result = System.Windows.MessageBox.Show(
            confirmMessage,
            confirmTitle,
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);

        if (result != MessageBoxResult.Yes)
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
                MessageBoxButton.OK,
                MessageBoxImage.Error);
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
        string newStyle = Path.Combine(appDirectory, "UninsWinTab.exe");
        if (File.Exists(newStyle))
            return newStyle;

        string legacyPreferred = Path.Combine(appDirectory, "unins000.exe");
        if (File.Exists(legacyPreferred))
            return legacyPreferred;

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
