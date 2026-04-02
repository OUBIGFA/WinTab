using WinTab.Core.Models;

namespace WinTab.App.Services;

internal static class ExplorerOpenVerbInterceptionPolicy
{
    public static bool NormalizeForNativeCurrentDirectoryBehavior(AppSettings settings)
    {
        _ = settings;
        return false;
    }

    public static bool ShouldEnableOpenVerbInterception(AppSettings settings, bool hasStableOpenVerbHandlerPath)
    {
        ArgumentNullException.ThrowIfNull(settings);

        return settings.EnableExplorerOpenVerbInterception &&
               settings.OpenChildFolderInNewTabFromActiveTab &&
               hasStableOpenVerbHandlerPath;
    }

    public static bool ShouldPersistAcrossReboot(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        return false;
    }
}
