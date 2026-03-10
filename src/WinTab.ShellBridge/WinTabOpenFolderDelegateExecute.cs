using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using WinTab.ShellBridge.Interop;
using WinTab.Platform.Win32;

namespace WinTab.ShellBridge;

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
[Guid(DelegateExecuteIds.OpenFolderDelegateExecuteClsid)]
public sealed class WinTabOpenFolderDelegateExecute : IExecuteCommand, IObjectWithSelection
{
    private const int S_OK = 0;
    private const int E_FAIL = unchecked((int)0x80004005);
    private const string TaskbarWindowClass = "Shell_TrayWnd";
    internal static Func<string, nint, bool, bool> SendOpenFolderRequest =
        static (target, foreground, allowRetry) => OpenRequestPipeClient.TrySendOpenFolderEx(target, foreground, allowRetry);
    internal static Action<string> OpenFallbackTarget = static target => TryOpenFallback(target);

    private IShellItemArray? _selection;
    private string? _parameters;

    public int SetKeyState(uint grfKeyState) => S_OK;

    public int SetParameters(string? pszParameters)
    {
        _parameters = pszParameters;
        return S_OK;
    }

    public int SetPosition(Point pt) => S_OK;

    public int SetShowWindow(int nShow) => S_OK;

    public int SetNoShowUI(int fNoShowUI) => S_OK;

    public int SetDirectory(string? pszDirectory) => S_OK;

    public int SetSelection(IShellItemArray? psia)
    {
        ReleaseSelection();
        _selection = psia;
        return S_OK;
    }

    public int GetSelection(ref Guid riid, out IntPtr ppv)
    {
        ppv = IntPtr.Zero;
        if (_selection is null)
            return E_FAIL;

        IntPtr unk = Marshal.GetIUnknownForObject(_selection);
        try
        {
            return Marshal.QueryInterface(unk, in riid, out ppv);
        }
        finally
        {
            Marshal.Release(unk);
        }
    }

    public int Execute()
    {
        try
        {
            string? rawTarget = TryGetTargetFromSelection();
            if (string.IsNullOrWhiteSpace(rawTarget))
                rawTarget = TryGetTargetFromParameters();

            if (!PathNormalization.TryNormalizeOpenTarget(rawTarget, out string target))
                return S_OK;

            OpenTargetInfo targetInfo = OpenTargetClassifier.Classify(target);
            if (targetInfo.RequiresNativeShellLaunch)
            {
                OpenFallbackTarget(target);
                return S_OK;
            }

            nint foreground = User32Native.GetForegroundWindow();
            bool allowRetry = !IsTaskbarForegroundWindow(foreground);
            if (SendOpenFolderRequest(target, foreground, allowRetry))
                return S_OK;

            OpenFallbackTarget(target);
            return S_OK;
        }
        finally
        {
            ReleaseSelection();
        }
    }

    private string? TryGetTargetFromParameters()
    {
        if (string.IsNullOrWhiteSpace(_parameters))
            return null;

        string value = _parameters.Trim();
        if (value.StartsWith("\"", StringComparison.Ordinal) && value.EndsWith("\"", StringComparison.Ordinal) && value.Length >= 2)
            return value[1..^1];

        return value;
    }

    private string? TryGetTargetFromSelection()
    {
        IShellItemArray? selection = _selection;
        if (selection is null)
            return null;

        if (selection.GetCount(out uint count) != S_OK || count == 0)
            return null;

        if (selection.GetItemAt(0, out IShellItem item) != S_OK || item is null)
            return null;

        try
        {
            string? fileSystemPath = TryGetDisplayName(item, Sigdn.FileSystemPath);
            if (!string.IsNullOrWhiteSpace(fileSystemPath))
                return fileSystemPath;

            return TryGetDisplayName(item, Sigdn.DesktopAbsoluteParsing);
        }
        finally
        {
            Marshal.FinalReleaseComObject(item);
        }
    }

    private static string? TryGetDisplayName(IShellItem item, Sigdn sigdn)
    {
        if (item.GetDisplayName(sigdn, out IntPtr rawName) != S_OK || rawName == IntPtr.Zero)
            return null;

        try
        {
            return Marshal.PtrToStringUni(rawName);
        }
        finally
        {
            Marshal.FreeCoTaskMem(rawName);
        }
    }

    private static void TryOpenFallback(string target)
    {
        if (NativeShellLauncher.TryOpen(target))
            return;

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{target}\"",
                UseShellExecute = true
            });
        }
        catch
        {
            // ignore
        }
    }

    private static bool IsTaskbarForegroundWindow(nint foregroundHwnd, Func<nint, string>? classNameResolver = null)
    {
        if (foregroundHwnd == 0)
            return false;

        string className = (classNameResolver ?? ResolveWindowClassName)(foregroundHwnd);
        return string.Equals(className, TaskbarWindowClass, StringComparison.OrdinalIgnoreCase);
    }

    private static string ResolveWindowClassName(nint hwnd)
    {
        var className = new StringBuilder(64);
        int written = User32Native.GetClassName(hwnd, className, className.Capacity);
        if (written <= 0)
            return string.Empty;

        return className.ToString();
    }

    private void ReleaseSelection()
    {
        if (_selection is null)
            return;

        Marshal.FinalReleaseComObject(_selection);
        _selection = null;
    }
}
