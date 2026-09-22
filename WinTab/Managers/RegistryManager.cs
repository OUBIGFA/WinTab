using System;
using Microsoft.Win32;
using WinTab.Helpers;

namespace WinTab.Managers;

public static class RegistryManager
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string StartupApprovedKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run";
    private const string ExplorerAdvancedKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced";
    private static readonly string? ExecutablePath = Helper.GetExecutablePath();
    private static readonly string? StartupCommand = string.IsNullOrWhiteSpace(ExecutablePath)
        ? null
        : $"\"{ExecutablePath}\" {Constants.BackgroundLaunchArg}";
    public static bool IsStartupEnabled => IsStartupEnabledUnder(Registry.CurrentUser);

    internal static bool IsStartupEnabledUnder(RegistryKey root) => IsInStartup(root) && IsStartupApprovedEnabled(root);

    public static void ToggleStartup() => ToggleStartup(Registry.CurrentUser);

    internal static void ToggleStartup(RegistryKey root)
    {
        if (IsStartupEnabledUnder(root))
            RemoveFromStartup(root);
        else
            AddToStartup(root);
    }

    private static bool IsInStartup(RegistryKey root)
    {
        if (string.IsNullOrWhiteSpace(ExecutablePath)) return false;

        // Check if the application exists in the Run registry key and has the correct executable location
        using var key = root.OpenSubKey(RunKeyPath, false);
        var value = key?.GetValue(Constants.AppName) as string;
        return string.Equals(value, StartupCommand, StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, ExecutablePath, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsStartupApprovedEnabled(RegistryKey root)
    {
        using var key = root.OpenSubKey(StartupApprovedKeyPath, false);
        var value = key?.GetValue(Constants.AppName) as byte[];
        // Check first byte parity (even = enabled, odd = disabled), null and empty also mean enabled.
        return value == null || value.Length == 0 || value[0] % 2 == 0;
    }

    private static void AddToStartup(RegistryKey root)
    {
        if (string.IsNullOrWhiteSpace(ExecutablePath) || string.IsNullOrWhiteSpace(StartupCommand)) return;

        // Add to Run registry key
        using var runKey = root.CreateSubKey(RunKeyPath, writable: true);
        runKey.SetValue(Constants.AppName, StartupCommand);

        // Create enabled entry in StartupApproved
        var enabledData = new byte[12];
        enabledData[0] = 0x02; // Even value for enabled

        using var approvedKey = root.CreateSubKey(StartupApprovedKeyPath, writable: true);
        approvedKey.SetValue(Constants.AppName, enabledData, RegistryValueKind.Binary);
    }

    private static void RemoveFromStartup(RegistryKey root)
    {
        // Remove from Run registry key
        using var runKey = root.OpenSubKey(RunKeyPath, true);
        runKey?.DeleteValue(Constants.AppName, false);

        // Remove from StartupApproved
        using var approvedKey = root.OpenSubKey(StartupApprovedKeyPath, true);
        approvedKey?.DeleteValue(Constants.AppName, false);
    }

    public static int GetDefaultExplorerLaunchId()
    {
        using var key = OpenCurrentUserKey(ExplorerAdvancedKeyPath, false);
        if (key == null) return 1;
        return key.GetValue("LaunchTo") as int? ?? 1;
    }

    private static RegistryKey? OpenCurrentUserKey(string name, bool writable) => Registry.CurrentUser.OpenSubKey(name, writable);
}
