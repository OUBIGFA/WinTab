using System.Runtime.InteropServices;

namespace WinTab.Interop;

/// <summary>The shell's taskbar list (CLSID_TaskbarList). Only the ITaskbarList methods WinTab uses are declared.</summary>
[ComImport]
[Guid("56FDF344-FD6D-11d0-958A-006097C9A090")]
[ClassInterface(ClassInterfaceType.None)]
internal class TaskbarList
{
}

[ComImport]
[Guid("56FDF342-FD6D-11d0-958A-006097C9A090")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface ITaskbarList
{
    [PreserveSig]
    int HrInit();

    [PreserveSig]
    int AddTab(nint hWnd);

    [PreserveSig]
    int DeleteTab(nint hWnd);

    [PreserveSig]
    int ActivateTab(nint hWnd);

    [PreserveSig]
    int SetActiveAlt(nint hWnd);
}
