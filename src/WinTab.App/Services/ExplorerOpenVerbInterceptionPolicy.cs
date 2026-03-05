using WinTab.Core.Models;

namespace WinTab.App.Services;

internal static class ExplorerOpenVerbInterceptionPolicy
{
    public static bool NormalizeForNativeCurrentDirectoryBehavior(AppSettings settings)
    {
        return false;
    }

    public static bool ShouldEnableOpenVerbInterception(AppSettings settings, bool hasStableOpenVerbHandlerPath)
    {
        return settings.EnableExplorerOpenVerbInterception && hasStableOpenVerbHandlerPath;
    }
}
