using System;
using System.Drawing;
using System.Runtime.InteropServices;
using WinTab.WinAPI;

namespace WinTab.Helpers;

public static class MouseSimulator
{
    private const uint MOUSEEVENTF_MOVE        = 0x0001;
    private const uint MOUSEEVENTF_LEFTDOWN    = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP      = 0x0004;
    private const uint MOUSEEVENTF_MIDDLEDOWN  = 0x0020;
    private const uint MOUSEEVENTF_MIDDLEUP    = 0x0040;
    private const uint MOUSEEVENTF_ABSOLUTE    = 0x8000;
    private const uint MOUSEEVENTF_VIRTUALDESK = 0x4000;

    private const int SM_XVIRTUALSCREEN  = 76;
    private const int SM_YVIRTUALSCREEN  = 77;
    private const int SM_CXVIRTUALSCREEN = 78;
    private const int SM_CYVIRTUALSCREEN = 79;

    /// <summary>
    /// Sends a middle-button click at the specified screen-space point.
    /// Windows 11 File Explorer's tab strip closes the tab under the cursor on a middle-click,
    /// so this is used as the close-tab primitive — it is dramatically faster than walking the
    /// UI Automation tree to find and invoke the close button. The synthesized events are flagged
    /// as injected and use absolute virtual-desktop coordinates so they reach the correct
    /// monitor without moving the visible cursor.
    /// </summary>
    public static void SendMiddleClick(Point screenPoint)
    {
        var (absX, absY) = ToAbsoluteVirtual(screenPoint);

        var inputs = new[]
        {
            CreateMouseInput(absX, absY, MOUSEEVENTF_MIDDLEDOWN | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK),
            CreateMouseInput(absX, absY, MOUSEEVENTF_MIDDLEUP   | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK),
        };

        WinApi.SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
    }

    private static (int absX, int absY) ToAbsoluteVirtual(Point screenPoint)
    {
        var x = WinApi.GetSystemMetrics(SM_XVIRTUALSCREEN);
        var y = WinApi.GetSystemMetrics(SM_YVIRTUALSCREEN);
        var w = Math.Max(1, WinApi.GetSystemMetrics(SM_CXVIRTUALSCREEN));
        var h = Math.Max(1, WinApi.GetSystemMetrics(SM_CYVIRTUALSCREEN));

        var absX = (int)(((long)(screenPoint.X - x) * 65535 + w / 2) / w);
        var absY = (int)(((long)(screenPoint.Y - y) * 65535 + h / 2) / h);
        return (absX, absY);
    }

    private static INPUT CreateMouseInput(int dx, int dy, uint flags)
    {
        return new INPUT
        {
            Type = InputType.Mouse,
            Data = new InputUnion
            {
                Mouse = new MOUSEINPUT
                {
                    dx = dx,
                    dy = dy,
                    mouseData = 0,
                    dwFlags = flags,
                    time = 0,
                    dwExtraInfo = 0,
                }
            }
        };
    }
}
