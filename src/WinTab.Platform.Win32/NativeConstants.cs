namespace WinTab.Platform.Win32;

/// <summary>
/// Win32 API constants organized by category.
/// </summary>
public static class NativeConstants
{
    // ─── Window Styles (WS_*) ───────────────────────────────────────────────

    public const long WS_OVERLAPPED       = 0x00000000L;
    public const long WS_POPUP            = 0x80000000L;
    public const long WS_CHILD            = 0x40000000L;
    public const long WS_MINIMIZE         = 0x20000000L;
    public const long WS_VISIBLE          = 0x10000000L;
    public const long WS_DISABLED         = 0x08000000L;
    public const long WS_CLIPSIBLINGS     = 0x04000000L;
    public const long WS_CLIPCHILDREN     = 0x02000000L;
    public const long WS_MAXIMIZE         = 0x01000000L;
    public const long WS_CAPTION          = 0x00C00000L;
    public const long WS_BORDER           = 0x00800000L;
    public const long WS_DLGFRAME         = 0x00400000L;
    public const long WS_VSCROLL          = 0x00200000L;
    public const long WS_HSCROLL          = 0x00100000L;
    public const long WS_SYSMENU          = 0x00080000L;
    public const long WS_THICKFRAME       = 0x00040000L;
    public const long WS_MINIMIZEBOX      = 0x00020000L;
    public const long WS_MAXIMIZEBOX      = 0x00010000L;
    public const long WS_OVERLAPPEDWINDOW = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU |
                                            WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX;

    // ─── Extended Window Styles (WS_EX_*) ───────────────────────────────────

    public const long WS_EX_DLGMODALFRAME   = 0x00000001L;
    public const long WS_EX_NOPARENTNOTIFY  = 0x00000004L;
    public const long WS_EX_TOPMOST         = 0x00000008L;
    public const long WS_EX_ACCEPTFILES     = 0x00000010L;
    public const long WS_EX_TRANSPARENT     = 0x00000020L;
    public const long WS_EX_MDICHILD        = 0x00000040L;
    public const long WS_EX_TOOLWINDOW      = 0x00000080L;
    public const long WS_EX_WINDOWEDGE      = 0x00000100L;
    public const long WS_EX_CLIENTEDGE      = 0x00000200L;
    public const long WS_EX_CONTEXTHELP     = 0x00000400L;
    public const long WS_EX_RIGHT           = 0x00001000L;
    public const long WS_EX_RTLREADING      = 0x00002000L;
    public const long WS_EX_CONTROLPARENT   = 0x00010000L;
    public const long WS_EX_STATICEDGE      = 0x00020000L;
    public const long WS_EX_APPWINDOW       = 0x00040000L;
    public const long WS_EX_LAYERED         = 0x00080000L;
    public const long WS_EX_NOINHERITLAYOUT = 0x00100000L;
    public const long WS_EX_NOREDIRECTIONBITMAP = 0x00200000L;
    public const long WS_EX_COMPOSITED      = 0x02000000L;
    public const long WS_EX_NOACTIVATE      = 0x08000000L;

    // ─── Show Window Commands (SW_*) ────────────────────────────────────────

    public const int SW_HIDE            = 0;
    public const int SW_SHOWNORMAL      = 1;
    public const int SW_SHOWMINIMIZED   = 2;
    public const int SW_SHOWMAXIMIZED   = 3;
    public const int SW_MAXIMIZE        = 3;
    public const int SW_SHOWNOACTIVATE  = 4;
    public const int SW_SHOW            = 5;
    public const int SW_MINIMIZE        = 6;
    public const int SW_SHOWMINNOACTIVE = 7;
    public const int SW_SHOWNA          = 8;
    public const int SW_RESTORE         = 9;
    public const int SW_SHOWDEFAULT     = 10;
    public const int SW_FORCEMINIMIZE   = 11;

    // ─── Set Window Pos Flags (SWP_*) ───────────────────────────────────────

    public const uint SWP_NOSIZE         = 0x0001;
    public const uint SWP_NOMOVE         = 0x0002;
    public const uint SWP_NOZORDER       = 0x0004;
    public const uint SWP_NOREDRAW       = 0x0008;
    public const uint SWP_NOACTIVATE     = 0x0010;
    public const uint SWP_FRAMECHANGED   = 0x0020;
    public const uint SWP_SHOWWINDOW     = 0x0040;
    public const uint SWP_HIDEWINDOW     = 0x0080;
    public const uint SWP_NOCOPYBITS     = 0x0100;
    public const uint SWP_NOOWNERZORDER  = 0x0200;
    public const uint SWP_NOSENDCHANGING = 0x0400;
    public const uint SWP_ASYNCWINDOWPOS = 0x4000;

    // ─── Window Messages (WM_*) ─────────────────────────────────────────────

    public const int WM_NULL             = 0x0000;
    public const int WM_DESTROY          = 0x0002;
    public const int WM_CLOSE            = 0x0010;
    public const int WM_QUIT             = 0x0012;
    public const int WM_GETICON          = 0x007F;
    public const int WM_NCHITTEST        = 0x0084;
    public const int WM_HOTKEY           = 0x0312;
    public const int WM_MOUSEMOVE        = 0x0200;
    public const int WM_LBUTTONDOWN      = 0x0201;
    public const int WM_LBUTTONUP        = 0x0202;
    public const int WM_RBUTTONDOWN      = 0x0204;
    public const int WM_RBUTTONUP        = 0x0205;
    public const int WM_NCLBUTTONDOWN    = 0x00A1;
    public const int WM_COMMAND          = 0x0111;
    public const int WM_SYSCOMMAND       = 0x0112;

    // Explorer internal command IDs
    public const int EXPLORER_CMD_OPEN_NEW_TAB = 0xA21B;
    public const int EXPLORER_CMD_CLOSE_TAB    = 0xA021;

    // ─── WM_GETICON wParam values ───────────────────────────────────────────

    public const int ICON_SMALL  = 0;
    public const int ICON_BIG    = 1;
    public const int ICON_SMALL2 = 2;

    // ─── WinEvent Constants ─────────────────────────────────────────────────

    public const uint EVENT_MIN                    = 0x00000001;
    public const uint EVENT_MAX                    = 0x7FFFFFFF;
    public const uint EVENT_SYSTEM_FOREGROUND      = 0x0003;
    public const uint EVENT_SYSTEM_MOVESIZESTART   = 0x000A;
    public const uint EVENT_SYSTEM_MOVESIZEEND     = 0x000B;
    public const uint EVENT_SYSTEM_MINIMIZESTART   = 0x0016;
    public const uint EVENT_SYSTEM_MINIMIZEEND     = 0x0017;
    public const uint EVENT_OBJECT_CREATE          = 0x8000;
    public const uint EVENT_OBJECT_DESTROY         = 0x8001;
    public const uint EVENT_OBJECT_SHOW            = 0x8002;
    public const uint EVENT_OBJECT_HIDE            = 0x8003;
    public const uint EVENT_OBJECT_LOCATIONCHANGE  = 0x800B;
    public const uint EVENT_OBJECT_NAMECHANGE      = 0x800C;

    public const uint WINEVENT_OUTOFCONTEXT   = 0x0000;
    public const uint WINEVENT_SKIPOWNPROCESS = 0x0002;

    // ─── ShellExecuteEx ─────────────────────────────────────────────────────
    public const uint SEE_MASK_INVOKEIDLIST = 0x0000000C;
    public const uint SEE_MASK_FLAG_NO_UI   = 0x00000400;
    public const uint SEE_MASK_ASYNCOK      = 0x00100000;

    // ─── DWM Window Attributes ──────────────────────────────────────────────

    public const int DWMWA_CLOAK = 13;
    public const int DWMWA_CLOAKED = 14;
    public const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
    public const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;
    public const int DWMWA_SYSTEMBACKDROP_TYPE     = 38;

    // DWM_WINDOW_CORNER_PREFERENCE values
    public const int DWMWCP_DEFAULT    = 0;
    public const int DWMWCP_DONOTROUND = 1;
    public const int DWMWCP_ROUND      = 2;
    public const int DWMWCP_ROUNDSMALL = 3;

    // DWM_SYSTEMBACKDROP_TYPE values
    public const int DWMSBT_AUTO       = 0;
    public const int DWMSBT_NONE       = 1;
    public const int DWMSBT_MAINWINDOW = 2; // Mica
    public const int DWMSBT_TRANSIENTWINDOW = 3; // Acrylic
    public const int DWMSBT_TABBEDWINDOW    = 4; // Tabbed

    // ─── Hotkey Modifiers (MOD_*) ───────────────────────────────────────────

    public const uint MOD_ALT      = 0x0001;
    public const uint MOD_CONTROL  = 0x0002;
    public const uint MOD_SHIFT    = 0x0004;
    public const uint MOD_WIN      = 0x0008;
    public const uint MOD_NOREPEAT = 0x4000;

    // ─── GetAncestor Flags ──────────────────────────────────────────────────

    public const uint GA_PARENT    = 1;
    public const uint GA_ROOT      = 2;
    public const uint GA_ROOTOWNER = 3;

    // ─── Object Identifiers ─────────────────────────────────────────────────

    public const int OBJID_WINDOW = 0x00000000;

    // ─── Monitor Info Flags ─────────────────────────────────────────────────

    public const int MONITOR_DEFAULTTONULL    = 0x00000000;
    public const int MONITOR_DEFAULTTOPRIMARY = 0x00000001;
    public const int MONITOR_DEFAULTTONEAREST = 0x00000002;

    // ─── GetWindowLong / SetWindowLong Indexes ──────────────────────────────

    public const int GWL_STYLE   = -16;
    public const int GWL_EXSTYLE = -20;
    public const int GWLP_HWNDPARENT = -8;
    public const int GWLP_WNDPROC    = -4;

    // ─── GetClassLongPtr Indexes ────────────────────────────────────────────

    public const int GCLP_HICONSM = -34;
    public const int GCLP_HICON   = -14;

    // ─── GetSystemMetrics Indexes ───────────────────────────────────────────

    public const int SM_CYCAPTION  = 4;
    public const int SM_CYFRAME    = 33;
    public const int SM_CXSCREEN   = 0;
    public const int SM_CYSCREEN   = 1;
    public const int SM_CMONITORS  = 80;

    // ─── Low-Level Mouse Hook ───────────────────────────────────────────────

    public const int WH_MOUSE_LL = 14;

    // ─── SetWindowPos HWND constants ────────────────────────────────────────

    public static readonly IntPtr HWND_TOP       = IntPtr.Zero;
    public static readonly IntPtr HWND_BOTTOM    = new(1);
    public static readonly IntPtr HWND_TOPMOST   = new(-1);
    public static readonly IntPtr HWND_NOTOPMOST = new(-2);
    public static readonly IntPtr HWND_MESSAGE   = new(-3);

    // ─── Process Access Rights ──────────────────────────────────────────────

    public const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;

    // ─── Foreground Activation ───────────────────────────────────────────────

    public const uint ASFW_ANY = 0xFFFFFFFF;

    // ─── SHGetFileInfo Flags ────────────────────────────────────────────────

    public const uint SHGFI_ICON      = 0x000000100;
    public const uint SHGFI_SMALLICON = 0x000000001;
    public const uint SHGFI_LARGEICON = 0x000000000;
}
