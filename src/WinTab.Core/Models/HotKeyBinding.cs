using WinTab.Core.Enums;

namespace WinTab.Core.Models;

public sealed class HotKeyBinding
{
    public HotKeyAction Action { get; set; }
    /// <summary>Win32 modifier flags: MOD_ALT=0x1, MOD_CONTROL=0x2, MOD_SHIFT=0x4, MOD_WIN=0x8</summary>
    public uint Modifiers { get; set; }
    /// <summary>Virtual key code.</summary>
    public uint Key { get; set; }
    public bool Enabled { get; set; } = true;

    public string DisplayString
    {
        get
        {
            var parts = new List<string>();
            if ((Modifiers & 0x0001) != 0) parts.Add("Alt");
            if ((Modifiers & 0x0002) != 0) parts.Add("Ctrl");
            if ((Modifiers & 0x0004) != 0) parts.Add("Shift");
            if ((Modifiers & 0x0008) != 0) parts.Add("Win");
            parts.Add($"0x{Key:X2}");
            return string.Join(" + ", parts);
        }
    }
}
