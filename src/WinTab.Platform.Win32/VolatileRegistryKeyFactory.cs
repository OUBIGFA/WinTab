using System.ComponentModel;
using Microsoft.Win32;

namespace WinTab.Platform.Win32;

public static class VolatileRegistryKeyFactory
{
    public static RegistryKey CreateCurrentUserVolatileSubKey(string subKeyPath)
    {
        return CreateCurrentUserVolatileSubKey(subKeyPath, RegistryView.Default);
    }

    public static RegistryKey CreateCurrentUserVolatileSubKey(string subKeyPath, RegistryView view)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(subKeyPath);

        using RegistryKey currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);

        int samDesired = NativeConstants.KEY_READ | NativeConstants.KEY_WRITE;
        if (view == RegistryView.Registry64)
            samDesired |= NativeConstants.KEY_WOW64_64KEY;
        else if (view == RegistryView.Registry32)
            samDesired |= NativeConstants.KEY_WOW64_32KEY;

        int result = NativeMethods.RegCreateKeyEx(
            currentUser.Handle,
            subKeyPath,
            0,
            null,
            NativeConstants.REG_OPTION_VOLATILE,
            samDesired,
            IntPtr.Zero,
            out var createdHandle,
            out _);

        if (result != 0)
            throw new Win32Exception(result, $"Failed to create volatile registry key: HKCU\\{subKeyPath}");

        return RegistryKey.FromHandle(createdHandle, view);
    }
}
