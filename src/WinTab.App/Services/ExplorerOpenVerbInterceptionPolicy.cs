using WinTab.Core.Models;

namespace WinTab.App.Services;

internal static class ExplorerOpenVerbInterceptionPolicy
{
    public static bool NormalizeForNativeCurrentDirectoryBehavior(AppSettings settings)
    {
        if (settings.OpenChildFolderInNewTabFromActiveTab)
            return false;

        if (!settings.EnableExplorerOpenVerbInterception)
            return false;

        settings.EnableExplorerOpenVerbInterception = false;
        return true;
    }

    public static bool ShouldEnableOpenVerbInterception(AppSettings settings, bool hasStableOpenVerbHandlerPath)
    {
        return settings.EnableExplorerOpenVerbInterception &&
               settings.OpenChildFolderInNewTabFromActiveTab &&
               hasStableOpenVerbHandlerPath;
    }
}
