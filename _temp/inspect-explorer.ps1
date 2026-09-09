Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;
public static class W {
  [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Unicode)] public static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string title);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder sb, int max);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr h, int idx);
  [DllImport("user32.dll")] public static extern bool GetLayeredWindowAttributes(IntPtr h, out uint key, out byte alpha, out uint flags);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  public delegate bool PropEnumProcEx(IntPtr hwnd, IntPtr lpszString, IntPtr hData, IntPtr dwData);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int EnumPropsEx(IntPtr hWnd, PropEnumProcEx lpEnumFunc, IntPtr lParam);
  public static List<string> Props(IntPtr h) {
    var list = new List<string>();
    EnumPropsEx(h, (hw, s, d, l) => { string name = (((long)s) >> 16) == 0 ? ("#" + (((long)s) & 0xFFFF)) : Marshal.PtrToStringUni(s); list.Add(name + "=" + d.ToString("X")); return true; }, IntPtr.Zero);
    return list;
  }
  public static List<IntPtr> FindAll(IntPtr parent, string cls) {
    var l = new List<IntPtr>(); IntPtr h = IntPtr.Zero;
    while ((h = FindWindowEx(parent, h, cls, null)) != IntPtr.Zero) l.Add(h);
    return l;
  }
}
"@
"foreground=0x{0:X}" -f [W]::GetForegroundWindow().ToInt64()
foreach ($w in [W]::FindAll([IntPtr]::Zero, "CabinetWClass")) {
  $sb = New-Object System.Text.StringBuilder 512; [void][W]::GetWindowText($w, $sb, 512)
  $procId = [uint32]0; $tid = [W]::GetWindowThreadProcessId($w, [ref]$procId)
  $ex = [W]::GetWindowLong($w, -20); $layered = ($ex -band 0x80000) -ne 0
  $k=[uint32]0;$a=[byte]0;$f=[uint32]0; $ok = [W]::GetLayeredWindowAttributes($w, [ref]$k, [ref]$a, [ref]$f)
  $tabs = [W]::FindAll($w, "ShellTabWindowClass")
  "HWND=0x{0:X} pid={1} tid={2} visible={3} iconic={4} layered={5} alpha={6} attrOk={7} tabs={8} title='{9}'" -f $w.ToInt64(), $procId, $tid, [W]::IsWindowVisible($w), [W]::IsIconic($w), $layered, $a, $ok, $tabs.Count, $sb.ToString()
  "  props: " + ([W]::Props($w) -join ", ")
  foreach ($t in $tabs) { $sb2 = New-Object System.Text.StringBuilder 512; [void][W]::GetWindowText($t, $sb2, 512); "  tab 0x{0:X} '{1}' props: {2}" -f $t.ToInt64(), $sb2.ToString(), ([W]::Props($t) -join ", ") }
}
"---- explorer processes ----"
Get-Process explorer | Select-Object Id, StartTime, @{n='Threads';e={$_.Threads.Count}}, HandleCount, @{n='WS_MB';e={[int]($_.WorkingSet64/1MB)}} | Format-Table -AutoSize | Out-String -Width 200
"---- wintab process ----"
Get-Process WinTab | Select-Object Id, StartTime, @{n='Threads';e={$_.Threads.Count}}, HandleCount, @{n='WS_MB';e={[int]($_.WorkingSet64/1MB)}}, @{n='CPUsec';e={[int]$_.CPU}}, Path | Format-List | Out-String -Width 200
