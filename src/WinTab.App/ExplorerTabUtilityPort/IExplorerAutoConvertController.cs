namespace WinTab.App.ExplorerTabUtilityPort;

public interface IExplorerAutoConvertController
{
    bool IsAutoConvertEnabled { get; }
    void SetAutoConvertEnabled(bool enabled);
}
