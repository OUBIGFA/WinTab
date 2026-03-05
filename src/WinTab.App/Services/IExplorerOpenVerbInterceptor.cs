namespace WinTab.App.Services;

public interface IExplorerOpenVerbInterceptor
{
    void StartupSelfCheck(bool settingEnabled);
    void EnableOrRepair();
    void DisableAndRestore();
}
