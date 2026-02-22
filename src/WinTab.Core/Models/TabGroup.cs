namespace WinTab.Core.Models;

public sealed class TabGroup
{
    public Guid Id { get; init; } = Guid.NewGuid();
    public string Name { get; set; } = "Default";
    public List<TabItem> Tabs { get; } = [];
    public IntPtr ActiveHandle { get; set; }
    public DateTime CreatedAt { get; init; } = DateTime.UtcNow;

    public double Left { get; set; }
    public double Top { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
}
