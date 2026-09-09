using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;
using WinTab.Hooks;

internal static class ExplorerDesktopFlowTests
{
    public static IEnumerable<(string Name, System.Func<Task> Body)> All()
    {
        yield return ("desktop folders use the v1.0.1 window-registration flow", UsesWindowRegistrationFlow);
    }

    private static Task UsesWindowRegistrationFlow()
    {
        var watcherType = typeof(ExplorerWatcher);
        var assembly = watcherType.Assembly;

        Check.That(assembly.GetType("WinTab.Hooks.ExplorerDesktopOpenHook", throwOnError: false) == null,
            "Desktop folder opening must not attach a DefaultVerb interceptor to Explorer.");
        Check.That(assembly.GetType("WinTab.Hooks.DesktopFolderOpenQueue", throwOnError: false) == null,
            "Desktop folder opening must not defer the native request through a separate queue.");
        Check.That(watcherType.GetMethod("HookDesktopFolderOpen", BindingFlags.Instance | BindingFlags.NonPublic) == null,
            "ExplorerWatcher must let the existing window-registration path own desktop folder merges.");

        return Task.CompletedTask;
    }
}
