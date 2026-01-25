using System;
using System.Collections.Generic;

namespace WinTab.Models;

internal sealed class WindowGroup
{
    public Guid Id { get; } = Guid.NewGuid();
    public string Name { get; set; } = string.Empty;
    public List<WindowEntry> Members { get; } = new();
    public DateTime CreatedAt { get; } = DateTime.UtcNow;
}


