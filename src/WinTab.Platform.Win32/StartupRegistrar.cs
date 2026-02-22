using Microsoft.Win32;

namespace WinTab.Platform.Win32;

/// <summary>
/// Manages application startup registration via the Windows Registry
/// (HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run).
/// </summary>
public sealed class StartupRegistrar
{
    private const string RunKeyPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";

    private readonly string _appName;
    private readonly string _executablePath;

    /// <summary>
    /// Creates a new <see cref="StartupRegistrar"/> instance.
    /// </summary>
    /// <param name="appName">Display name used as the registry value name.</param>
    /// <param name="executablePath">Full path to the application executable.</param>
    public StartupRegistrar(string appName, string executablePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(appName);
        ArgumentException.ThrowIfNullOrWhiteSpace(executablePath);

        _appName = appName;
        _executablePath = executablePath;
    }

    /// <summary>
    /// Checks whether the application is currently registered to run at startup.
    /// </summary>
    /// <returns>True if the registry entry exists and points to the correct path.</returns>
    public bool IsEnabled()
    {
        try
        {
            using RegistryKey? key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
            if (key is null)
                return false;

            object? value = key.GetValue(_appName);
            if (value is not string path)
                return false;

            // Compare paths case-insensitively (Windows file system is case-insensitive).
            return string.Equals(
                NormalizePath(path),
                NormalizePath($"\"{_executablePath}\""),
                StringComparison.OrdinalIgnoreCase) ||
                string.Equals(
                NormalizePath(path),
                NormalizePath(_executablePath),
                StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception)
        {
            return false;
        }
    }

    /// <summary>
    /// Enable or disable startup registration.
    /// </summary>
    /// <param name="enable">True to add the registry entry; false to remove it.</param>
    public void SetEnabled(bool enable)
    {
        try
        {
            using RegistryKey? key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true);
            if (key is null)
                return;

            if (enable)
            {
                // Quote the path to handle spaces in the executable path.
                key.SetValue(_appName, $"\"{_executablePath}\"");
            }
            else
            {
                key.DeleteValue(_appName, throwOnMissingValue: false);
            }
        }
        catch (Exception)
        {
            // Silently ignore permission or registry access errors.
            // The caller should check IsEnabled() to verify the result.
        }
    }

    /// <summary>
    /// Normalizes a file path for comparison: trims quotes and whitespace.
    /// </summary>
    private static string NormalizePath(string path)
    {
        return path.Trim().Trim('"').Trim();
    }
}
