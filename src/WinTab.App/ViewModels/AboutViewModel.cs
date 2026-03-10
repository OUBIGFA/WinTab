using System.Diagnostics;
using System.IO;
using System.Reflection;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.UI.Localization;

namespace WinTab.App.ViewModels;

public partial class AboutViewModel : ObservableObject
{
    private const string RepositoryUrl = "https://github.com/OUBIGFA/WinTab";

    private readonly Logger _logger;

    [ObservableProperty]
    private string _version;

    [ObservableProperty]
    private string _logPath;

    [ObservableProperty]
    private string _modeText;

    public AboutViewModel(Logger logger)
    {
        _logger = logger;

        // Get version from assembly
        var assembly = Assembly.GetEntryAssembly() ?? Assembly.GetExecutingAssembly();
        var version = assembly.GetName().Version;
        _version = version is not null ? $"{version.Major}.{version.Minor}.{version.Build}" : "0.0.0";

        _logPath = AppPaths.LogPath;
        _modeText = LocalizationManager.GetString(AppPaths.IsPortable ? "About_ModePortable" : "About_ModeInstalled");
    }

    [RelayCommand]
    private void OpenLog()
    {
        try
        {
            if (File.Exists(LogPath))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = LogPath,
                    UseShellExecute = true
                });
            }
            else
            {
                _logger.Warn($"Log file not found: {LogPath}");
            }
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or FileNotFoundException)
        {
            _logger.Error("Failed to open log file.", ex);
        }
    }

    [RelayCommand]
    private void OpenLogFolder()
    {
        try
        {
            string? directory = Path.GetDirectoryName(LogPath);
            if (directory is not null && Directory.Exists(directory))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = directory,
                    UseShellExecute = true
                });
            }
            else
            {
                _logger.Warn($"Log directory not found: {directory}");
            }
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or FileNotFoundException)
        {
            _logger.Error("Failed to open log folder.", ex);
        }
    }

    [RelayCommand]
    private void OpenGitHub()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = RepositoryUrl,
                UseShellExecute = true
            });
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or FileNotFoundException)
        {
            _logger.Error("Failed to open project repository.", ex);
        }
    }
}
