using System.Threading.Tasks;
using WinTab.Models;

namespace WinTab.Hooks;

public partial class ExplorerWatcher
{
    /// <summary>External launchers identify the requested page explicitly, without first opening another frame.</summary>
    public async Task<bool> OpenRecycleBinAsync()
    {
        var opened = false;
        await RunShellWorkAsync(async () =>
        {
            if (!CanReuseDesktopFolder) return;
            opened = await OpenTabNavigateWithSelection(
                new WindowRecord("shell:::{645FF040-5081-101B-9F08-00AA002F954E}"));
        });
        return opened;
    }
}
