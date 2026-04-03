using WinTab.Core.Models;

namespace WinTab.App.Services;

public interface IExplorerOpenVerbConfigurationController
{
    void ReconfigureForCurrentSettings(AppSettings settings);
}
