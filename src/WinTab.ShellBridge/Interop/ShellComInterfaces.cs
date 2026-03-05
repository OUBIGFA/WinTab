using System.Runtime.InteropServices;

namespace WinTab.ShellBridge.Interop;

[StructLayout(LayoutKind.Sequential)]
public struct Point
{
    public int X;
    public int Y;
}

public enum Sigdn : uint
{
    FileSystemPath = 0x80058000,
    DesktopAbsoluteParsing = 0x80028000
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
public interface IShellItem
{
    [PreserveSig]
    int BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);

    [PreserveSig]
    int GetParent(out IShellItem ppsi);

    [PreserveSig]
    int GetDisplayName(Sigdn sigdnName, out IntPtr ppszName);

    [PreserveSig]
    int GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);

    [PreserveSig]
    int Compare(IShellItem psi, uint hint, out int piOrder);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("b63ea76d-1f85-456f-a19c-48159efa858b")]
public interface IShellItemArray
{
    [PreserveSig]
    int BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppvOut);

    [PreserveSig]
    int GetPropertyStore(uint flags, ref Guid riid, out IntPtr ppv);

    [PreserveSig]
    int GetPropertyDescriptionList(ref Guid keyType, ref Guid riid, out IntPtr ppv);

    [PreserveSig]
    int GetAttributes(uint attribFlags, uint sfgaoMask, out uint psfgaoAttribs);

    [PreserveSig]
    int GetCount(out uint pdwNumItems);

    [PreserveSig]
    int GetItemAt(uint dwIndex, out IShellItem ppsi);

    [PreserveSig]
    int EnumItems(out IntPtr ppenumShellItems);
}

[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("7F9185B0-CB92-43c5-80A9-92277A4F7B54")]
public interface IExecuteCommand
{
    [PreserveSig]
    int SetKeyState(uint grfKeyState);

    [PreserveSig]
    int SetParameters([MarshalAs(UnmanagedType.LPWStr)] string? pszParameters);

    [PreserveSig]
    int SetPosition(Point pt);

    [PreserveSig]
    int SetShowWindow(int nShow);

    [PreserveSig]
    int SetNoShowUI(int fNoShowUI);

    [PreserveSig]
    int SetDirectory([MarshalAs(UnmanagedType.LPWStr)] string? pszDirectory);

    [PreserveSig]
    int Execute();
}

[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("1c9cd5bb-98e9-4491-a60f-31aacc72b83c")]
public interface IObjectWithSelection
{
    [PreserveSig]
    int SetSelection([MarshalAs(UnmanagedType.Interface)] IShellItemArray? psia);

    [PreserveSig]
    int GetSelection(ref Guid riid, out IntPtr ppv);
}
