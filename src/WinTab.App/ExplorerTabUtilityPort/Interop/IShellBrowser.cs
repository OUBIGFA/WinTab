// This file is derived from ExplorerTabUtility (MIT License).
// Source: E:\_BIGFA Free\_code\ExplorerTabUtility

using System.Runtime.InteropServices;

namespace WinTab.App.ExplorerTabUtilityPort.Interop;

[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("000214E2-0000-0000-C000-000000000046")]
[ComImport]
internal interface IShellBrowser
{
    [PreserveSig]
    int GetWindow(out nint handle);
}
