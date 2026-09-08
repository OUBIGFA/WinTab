using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.Models;

internal static class ExplorerReuseSelectionTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("reusing a tab applies the requested file selection", () => ReuseSelectionAsync(["target.txt"], ["target.txt"]));
        yield return ("reusing a tab skips missing files without retaining an old selection", () => ReuseSelectionAsync(["missing.txt", "first.txt", "second.txt"], ["first.txt", "second.txt"]));
        yield return ("reusing a tab without a selection request preserves the current selection", () => ReuseSelectionAsync(null, ["previous.txt"]));
        yield return ("reuse does not report success when the selection view is unavailable", () => ReuseSelectionAsync(["target.txt"], ["previous.txt"], expectedReused: false, viewAvailable: false));
        yield return ("reuse does not report success when all requested files are missing", () => ReuseSelectionAsync(["missing.txt"], ["previous.txt"], expectedReused: false));
    }

    private static async Task ReuseSelectionAsync(string[]? requested, string[] expected, bool expectedReused = true, bool viewAvailable = true)
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(async () =>
        {
            using var lifetime = new CancellationTokenSource();
            using var openLock = new SemaphoreSlim(1);
            using var fixture = new ExplorerTabActivationTests.ActivationWindow();
            fixture.SetActive(0);
            var watcher = ExplorerTabActivationTests.CreateSelectionWatcher(lifetime);
            SetField(watcher, "_toOpenWindowsLock", openLock);
            SetField(watcher, "_windowEntryDictLock", new object());
            SetField(watcher, "_reuseTabs", true);

            var selected = new List<string> { "previous.txt" };
            var view = viewAvailable ? CreateFolderView(selected) : null;
            var dictionaryField = typeof(ExplorerWatcher).GetField("_windowEntryDict", BindingFlags.Instance | BindingFlags.NonPublic)!;
            var dictionary = Activator.CreateInstance(dictionaryField.FieldType)!;
            var browserType = dictionaryField.FieldType.GetGenericArguments()[0];
            var browser = ShellDispatchStub.Create(browserType, (method, arguments) => method switch
            {
                "get_HWND" => (int)fixture.Handle,
                "get_LocationURL" => "file:///C:/WinTab-reuse",
                "get_Document" => view,
                _ => throw new InvalidOperationException("Unexpected browser call: " + method)
            });
            var location = Helper.NormalizeLocation("C:/WinTab-reuse");
            var info = new WindowInfo
            {
                Identity = WindowIdentity.Capture(fixture.Handle),
                TabIdentity = WindowIdentity.Capture(fixture.FirstTab),
                HookedTopLevelHWnd = fixture.Handle,
                Location = location
            };
            dictionaryField.FieldType.GetMethods().Single(method => method.Name == "Add" && method.GetParameters().Length == 3)
                .Invoke(dictionary, [browser, info, (nint?)fixture.FirstTab]);
            dictionaryField.SetValue(watcher, dictionary);
            var searchMethod = typeof(ExplorerWatcher).GetMethod("TrySearchForTab", BindingFlags.Instance | BindingFlags.NonPublic)!;
            object?[] searchArguments = [location, (nint)0, (nint)0, null];
            var found = (bool)searchMethod.Invoke(watcher, searchArguments)!;
            Check.That(found && (nint)searchArguments[2]! == fixture.FirstTab,
                "The reuse search must resolve the isolated test tab before the full flow is allowed to run.");
            var openMethod = typeof(ExplorerWatcher).GetMethod("OpenTabNavigateWithSelection", BindingFlags.Instance | BindingFlags.NonPublic)!;

            var reused = await (Task<bool>)openMethod.Invoke(watcher,
                [new WindowRecord(location, selectedItems: requested), fixture.Handle])!;

            Check.Equal(expectedReused, reused, "Reuse must not claim success when the requested file selection cannot be applied.");
            Check.That(selected.SequenceEqual(expected), "The reused tab must contain the requested selection, without unrelated old files.");
            Check.Equal(2, ExplorerWindowDiscovery.GetAllExplorerTabs(fixture.Handle).Count(), "Reuse must not create another tab.");
        }, CancellationToken.None, TaskCreationOptions.None, scheduler).Unwrap();
    }

    private static object CreateFolderView(List<string> selected)
    {
        var viewType = typeof(ExplorerWatcher).Assembly.GetType("Shell32.ShellFolderView", throwOnError: true)!;
        var folderType = viewType.GetInterfaces().Append(viewType).SelectMany(type => type.GetProperties())
            .First(property => property.Name == "Folder").PropertyType;
        var itemType = folderType.GetInterfaces().Append(folderType).SelectMany(type => type.GetMethods())
            .First(method => method.Name == "ParseName").ReturnType;
        var parsedNames = new Dictionary<object, string>();
        var folder = ShellDispatchStub.Create(folderType, (method, arguments) =>
        {
            if (method != "ParseName")
                throw new InvalidOperationException("Unexpected folder call: " + method);
            var name = (string)arguments[0]!;
            if (name == "missing.txt")
                return null;
            var item = ShellDispatchStub.Create(itemType, (itemMethod, itemArguments) =>
                itemMethod == "get_Name" ? name : throw new InvalidOperationException("Unexpected item call: " + itemMethod));
            parsedNames.Add(item, name);
            return item;
        });
        return ShellDispatchStub.Create(viewType, (method, arguments) =>
        {
            if (method == "get_Folder")
                return folder;
            if (method != "SelectItem")
                throw new InvalidOperationException("Unexpected view call: " + method);
            var flags = Convert.ToInt32(arguments[1]);
            if ((flags & 4) != 0)
                selected.Clear();
            if ((flags & 1) != 0)
                selected.Add(parsedNames[arguments[0]!]);
            return null;
        });
    }

    private static void SetField(ExplorerWatcher watcher, string name, object value) =>
        typeof(ExplorerWatcher).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(watcher, value);
}

public class ShellDispatchStub : DispatchProxy
{
    private Func<string, object?[], object?> _invoke = null!;

    public static object Create(Type interfaceType, Func<string, object?[], object?> invoke)
    {
        var proxy = DispatchProxy.Create(interfaceType, typeof(ShellDispatchStub));
        ((ShellDispatchStub)proxy)._invoke = invoke;
        return proxy;
    }

    protected override object? Invoke(MethodInfo? targetMethod, object?[]? arguments) =>
        _invoke(targetMethod!.Name, arguments ?? []);
}
