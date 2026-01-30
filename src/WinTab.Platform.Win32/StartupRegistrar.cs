using Microsoft.Win32;

namespace WinTab.Platform.Win32;

public sealed class StartupRegistrar
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private readonly string _appName;
    private readonly string _executablePath;

    public StartupRegistrar(string appName, string executablePath)
    {
        _appName = appName;
        _executablePath = executablePath;
    }

    public bool IsEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, false);
        var value = key?.GetValue(_appName) as string;
        return !string.IsNullOrWhiteSpace(value);
    }

    public void SetEnabled(bool enable)
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, true) ??
                        Registry.CurrentUser.CreateSubKey(RunKeyPath);

        if (key is null)
        {
            return;
        }

        if (enable)
        {
            key.SetValue(_appName, $"\"{_executablePath}\"");
        }
        else
        {
            key.DeleteValue(_appName, false);
        }
    }
}
