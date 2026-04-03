using WinTab.Core.Models;

namespace WinTab.App.Services;

internal static class ExplorerOpenVerbInterceptionPolicy
{
    public static bool NormalizeForNativeCurrentDirectoryBehavior(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        bool expectedInterceptionState = settings.EnableAutoConvertExplorerWindows;
        if (settings.EnableExplorerOpenVerbInterception == expectedInterceptionState)
            return false;

        settings.EnableExplorerOpenVerbInterception = expectedInterceptionState;
        return true;
    }

    public static bool ShouldEnableOpenVerbInterception(AppSettings settings, bool hasStableOpenVerbHandlerPath)
    {
        ArgumentNullException.ThrowIfNull(settings);

        return settings.EnableAutoConvertExplorerWindows &&
               hasStableOpenVerbHandlerPath;
    }

    public static bool ShouldPersistAcrossReboot(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        return false;
    }
}
