namespace WinTab.App.Services;

public interface IExplorerOpenVerbInterceptor
{
    void StartupSelfCheck(bool settingEnabled, bool persistAcrossReboot);
    void EnableOrRepair(bool persistAcrossReboot);
    void DisableAndRestore(bool deleteBackup = true);
}
