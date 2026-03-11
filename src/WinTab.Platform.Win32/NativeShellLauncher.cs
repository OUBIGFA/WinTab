using System.Runtime.InteropServices;

namespace WinTab.Platform.Win32;

public static class NativeShellLauncher
{
    private const int S_OK = 0;
    internal static Func<string, nint> ParseDisplayNameToPidl = static target =>
    {
        int hr = NativeMethods.SHParseDisplayName(target, IntPtr.Zero, out IntPtr pidl, 0, out _);
        return hr == S_OK ? pidl : IntPtr.Zero;
    };
    internal static Func<nint, bool> OpenFolderByPidl =
        static pidl => NativeMethods.SHOpenFolderAndSelectItems(pidl, 0, IntPtr.Zero, 0) == S_OK;
    internal static Action<nint> ReleasePidl = static pidl =>
    {
        if (pidl != IntPtr.Zero)
            Marshal.FreeCoTaskMem(pidl);
    };

    public static bool TryOpen(string? target)
    {
        OpenTargetInfo targetInfo = OpenTargetClassifier.Classify(target);
        if (!targetInfo.IsShellNamespace)
            return false;

        nint pidl = IntPtr.Zero;

        try
        {
            pidl = TryParseTargetToPidl(targetInfo.NormalizedTarget);
            if (pidl == IntPtr.Zero)
                return false;

            return OpenFolderByPidl(pidl);
        }
        finally
        {
            if (pidl != IntPtr.Zero)
                ReleasePidl(pidl);
        }
    }

    private static nint TryParseTargetToPidl(string target)
    {
        foreach (string candidate in ShellNamespacePath.BuildNamespaceCandidates(target))
        {
            nint pidl = ParseDisplayNameToPidl(candidate);
            if (pidl != IntPtr.Zero)
                return pidl;
        }

        return IntPtr.Zero;
    }
}
