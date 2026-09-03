using System.Runtime.InteropServices;
using WinTab.WinAPI;

namespace WinTab.Helpers;

public static class KeyboardSimulator
{
    public static bool IsKeyPressed(int keyCode)
    {
        return (WinApi.GetAsyncKeyState(keyCode) & 0x8000) != 0;
    }

    public static void SendKeyPress(VirtualKey keyCode)
    {
        var inputs = new[] { CreateKeyInput(keyCode, KeyEventFlags.KeyDown), CreateKeyInput(keyCode, KeyEventFlags.KeyUp) };
        WinApi.SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
    }

    private static INPUT CreateKeyInput(VirtualKey keyCode, KeyEventFlags flags)
    {
        var wScan = (ushort)(WinApi.MapVirtualKey((uint)keyCode, 0) & 0xFFU);
        return new INPUT
        {
            Type = InputType.Keyboard,
            Data = new InputUnion
            {
                Keyboard = new KEYBDINPUT
                {
                    wVk = keyCode,
                    wScan = wScan,
                    dwFlags = flags
                }
            }
        };
    }
}
