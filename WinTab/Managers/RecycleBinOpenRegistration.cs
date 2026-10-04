using System;
using Microsoft.Win32;
using WinTab.Helpers;
using WinTab.Hooks;

namespace WinTab.Managers;

internal static class RecycleBinOpenRegistration
{
    private const string Page = @"Software\Classes\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}";
    private const string Shell = Page + @"\shell";
    private const string Open = Shell + @"\open";
    private const string Command = Open + @"\command";
    private const string Owner = "WinTab.RecycleBinOpen.v1";
    private static readonly string[] Metadata = ["WinTab.Owner", "WinTab.Command", "WinTab.HadDefault",
        "WinTab.Default", "WinTab.DefaultKind", "WinTab.HadPage", "WinTab.HadShell"];

    /// <summary>Shell associations are redirected by caller bitness; RocketDock needs the 32-bit view too.</summary>
    internal static void Update(bool enabled)
    {
        var executable = Helper.GetExecutablePath();
        if (string.IsNullOrWhiteSpace(executable)) return;
        foreach (var view in Environment.Is64BitOperatingSystem
            ? new[] { RegistryView.Registry64, RegistryView.Registry32 } : [RegistryView.Registry32])
        {
            try
            {
                using var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
                if (!Update(root, enabled, executable))
                    ExplorerDebugLog.Write($"Recycle Bin compatibility left custom registration unchanged view={view} enabled={enabled}");
            }
            catch (Exception exception) when (exception is UnauthorizedAccessException or System.Security.SecurityException or System.IO.IOException)
            {
                ExplorerDebugLog.Write($"Recycle Bin compatibility registration failed view={view} enabled={enabled}: {exception.Message}");
            }
        }
    }

    /// <summary>Only our own verb is changed; its original default survives process crashes and upgrades.</summary>
    internal static bool Update(RegistryKey root, bool enabled, string executablePath)
    {
        using var existing = root.OpenSubKey(Open);
        if (existing != null)
        {
            if (existing.GetValue("WinTab.Owner") as string != Owner) return !enabled;
            var savedCommand = existing.GetValue("WinTab.Command") as string;
            using var currentCommand = root.OpenSubKey(Command);
            var actualCommand = currentCommand?.GetValue("") as string;
            // A later customization wins, even when it retains our ownership marker.
            if (actualCommand != null && (actualCommand != savedCommand || currentCommand?.GetValue("DelegateExecute") as string != ""))
                return false;
            if (enabled && actualCommand != null)
            {
                using var currentShell = root.OpenSubKey(Shell);
                if (currentShell?.GetValue("") as string != "open") return false;
                var desired = BuildCommand(executablePath);
                if (actualCommand == desired) return true;
            }
            RemoveOwned(root, existing);
        }
        if (!enabled) return true;

        using var page = root.OpenSubKey(Page);
        using var shell = root.OpenSubKey(Shell);
        var previous = shell?.GetValue("", null, RegistryValueOptions.DoNotExpandEnvironmentNames);
        if (previous != null && previous is not string) return false;
        using (var verb = root.CreateSubKey(Open))
        {
            verb.SetValue("WinTab.HadPage", page != null ? 1 : 0);
            verb.SetValue("WinTab.HadShell", shell != null ? 1 : 0);
            verb.SetValue("WinTab.HadDefault", previous != null ? 1 : 0);
            verb.SetValue("WinTab.Default", previous as string ?? "");
            verb.SetValue("WinTab.DefaultKind", previous != null ? (int)shell!.GetValueKind("") : (int)RegistryValueKind.String);
            verb.SetValue("WinTab.Command", BuildCommand(executablePath));
            verb.SetValue("WinTab.Owner", Owner);
            verb.Flush(); // Persist recovery metadata before changing any default action.
        }
        using (var command = root.CreateSubKey(Command))
        {
            command.SetValue("", BuildCommand(executablePath));
            command.SetValue("DelegateExecute", "");
        }
        using (var writableShell = root.CreateSubKey(Shell)) writableShell.SetValue("", "open");
        return true;
    }

    private static string BuildCommand(string executablePath) => $"\"{executablePath}\" {Constants.OpenRecycleBinArg}";

    private static void RemoveOwned(RegistryKey root, RegistryKey metadata)
    {
        var hadShell = metadata.GetValue("WinTab.HadShell") is int shellFlag && shellFlag != 0;
        var hadPage = metadata.GetValue("WinTab.HadPage") is int pageFlag && pageFlag != 0;
        using (var shell = root.OpenSubKey(Shell, true))
        {
            if (shell?.GetValue("") as string == "open")
            {
                if (metadata.GetValue("WinTab.HadDefault") is int flag && flag != 0)
                    shell.SetValue("", metadata.GetValue("WinTab.Default") as string ?? "",
                        (RegistryValueKind)(int)metadata.GetValue("WinTab.DefaultKind", (int)RegistryValueKind.String)!);
                else shell.DeleteValue("", false);
            }
        }
        using (var command = root.OpenSubKey(Command, true))
        {
            command?.DeleteValue("", false);
            command?.DeleteValue("DelegateExecute", false);
        }
        using (var verb = root.OpenSubKey(Open, true))
            foreach (var name in Metadata) verb?.DeleteValue(name, false);
        DeleteIfEmpty(root, Command);
        DeleteIfEmpty(root, Open);
        if (!hadShell) DeleteIfEmpty(root, Shell);
        if (!hadPage) DeleteIfEmpty(root, Page);
    }

    private static void DeleteIfEmpty(RegistryKey root, string path)
    {
        using (var key = root.OpenSubKey(path))
            if (key == null || key.ValueCount != 0 || key.SubKeyCount != 0) return;
        root.DeleteSubKey(path, false);
    }
}
