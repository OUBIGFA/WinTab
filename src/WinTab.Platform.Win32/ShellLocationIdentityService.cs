using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace WinTab.Platform.Win32;

public sealed class ShellLocationIdentityService
{
    private const string ShellPrefix = "shell:::";

    public string NormalizeLocation(string? location)
    {
        if (string.IsNullOrWhiteSpace(location))
            return string.Empty;

        string value = location.Trim();

        if (value.IndexOf('%') >= 0)
            value = Environment.ExpandEnvironmentVariables(value);

        if (Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) && uri.IsFile)
            value = Uri.UnescapeDataString(uri.LocalPath);

        value = value.Trim(' ', '\'', '"', '\n', '\r', '\t');

        if (value.StartsWith("::", StringComparison.Ordinal))
            value = ShellPrefix + value[2..];
        else if (value.StartsWith("shell::", StringComparison.OrdinalIgnoreCase) &&
                 !value.StartsWith(ShellPrefix, StringComparison.OrdinalIgnoreCase))
            value = ShellPrefix + value[7..];
        else if (value.StartsWith("{", StringComparison.Ordinal) && value.EndsWith("}", StringComparison.Ordinal))
            value = ShellPrefix + value;

        value = value.Replace('/', '\\');

        if (IsFileSystemPathLike(value))
        {
            try
            {
                string full = Path.GetFullPath(value);
                return NormalizePath(full);
            }
            catch
            {
                return NormalizePath(value);
            }
        }

        return value.Normalize(NormalizationForm.FormKC);
    }

    public bool AreEquivalent(string? left, string? right)
    {
        string normalizedLeft = NormalizeLocation(left);
        string normalizedRight = NormalizeLocation(right);

        if (string.IsNullOrEmpty(normalizedLeft) || string.IsNullOrEmpty(normalizedRight))
            return false;

        if (string.Equals(normalizedLeft, normalizedRight, StringComparison.OrdinalIgnoreCase))
            return true;

        bool shellCompared = TryCompareByPidl(normalizedLeft, normalizedRight, out bool shellEqual);
        if (shellCompared)
            return shellEqual;

        return string.Equals(normalizedLeft, normalizedRight, StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryCompareByPidl(string left, string right, out bool equal)
    {
        equal = false;

        IntPtr leftPidl = IntPtr.Zero;
        IntPtr rightPidl = IntPtr.Zero;

        try
        {
            if (!TryParseDisplayName(left, out leftPidl) || !TryParseDisplayName(right, out rightPidl))
                return false;

            equal = NativeMethods.ILIsEqual(leftPidl, rightPidl);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (leftPidl != IntPtr.Zero)
                Marshal.FreeCoTaskMem(leftPidl);

            if (rightPidl != IntPtr.Zero)
                Marshal.FreeCoTaskMem(rightPidl);
        }
    }

    private static bool TryParseDisplayName(string value, out IntPtr pidl)
    {
        pidl = IntPtr.Zero;
        uint attrs = 0;
        int hr = NativeMethods.SHParseDisplayName(value, IntPtr.Zero, out pidl, 0, out attrs);
        return hr == 0 && pidl != IntPtr.Zero;
    }

    private static bool IsFileSystemPathLike(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return false;

        if (value.StartsWith("\\\\", StringComparison.Ordinal))
            return true;

        return value.Length >= 3 &&
               char.IsLetter(value[0]) &&
               value[1] == ':' &&
               (value[2] == '\\' || value[2] == '/');
    }

    private static string NormalizePath(string value)
    {
        string normalized = value
            .Replace('/', '\\')
            .TrimEnd('\\')
            .Normalize(NormalizationForm.FormKC);

        if (normalized.Length == 2 && normalized[1] == ':')
            return normalized + "\\";

        return normalized;
    }
}
