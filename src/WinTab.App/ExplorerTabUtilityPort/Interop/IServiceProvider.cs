// This file is derived from ExplorerTabUtility (MIT License).
// Source: E:\_BIGFA Free\_code\ExplorerTabUtility

using System.Runtime.InteropServices;

namespace WinTab.App.ExplorerTabUtilityPort.Interop;

[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("6d5140c1-7436-11ce-8034-00aa006009fa")]
[ComImport]
internal interface IServiceProvider
{
    [PreserveSig]
    int QueryService(ref Guid guidService, ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out IShellBrowser? ppvObject);
}
