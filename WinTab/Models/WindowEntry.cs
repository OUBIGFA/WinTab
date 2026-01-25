using System;
using WinTab.Hooks;

namespace WinTab.Models;

public sealed class WindowEntry
{
    public nint Handle { get; set; }
    public string Title { get; set; } = string.Empty;
    public string ProcessPath { get; set; } = string.Empty;
    public WindowHostType HostType { get; set; } = WindowHostType.Unknown;
    public long CreatedAt { get; } = Environment.TickCount;
    public WindowLifecycleState State { get; set; } = WindowLifecycleState.Unknown;
    public DateTime? LastActivatedAt { get; set; }
}


