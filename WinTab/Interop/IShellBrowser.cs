using System.Runtime.InteropServices;

namespace WinTab.Interop;

[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("000214E2-0000-0000-C000-000000000046")]
[ComImport]
public interface IShellBrowser
{
    [PreserveSig]
    int GetWindow(out nint handle);

    // Skip ContextSensitiveHelp, InsertMenusSB, SetMenuSB, RemoveMenusSB, SetStatusTextSB, EnableModelessSB
    // and TranslateAcceleratorSB (7 methods).
    void _VtblGap1_7();

    /// <summary>Browses to the item in this window, or in a new browser depending on the SBSP flags.</summary>
    [PreserveSig]
    int BrowseObject(nint pidl, uint flags);
}
