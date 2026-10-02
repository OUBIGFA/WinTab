using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.Interop;
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
        yield return ("file-location reuse preserves a Start menu shortcut selection", () => ReuseSelectionAsync(null, ["测试应用.lnk"],
            capturedItems: [("测试应用", @"C:\WinTab-reuse\测试应用.lnk", true)]));
        yield return ("file-location reuse preserves hidden extensions in a mixed selection", () => ReuseSelectionAsync(null, ["report.txt", "target.txt"],
            capturedItems: [("report", @"C:\WinTab-reuse\report.txt", true), ("target.txt", @"C:\WinTab-reuse\target.txt", true)]));
        yield return ("file-location reuse retains virtual item names", () => ReuseSelectionAsync(null, ["Control Panel"],
            capturedItems: [("Control Panel", "::{26EE0668-A00A-44D7-9371-BEB064C98683}", false)]));
        yield return ("file-location reuse retains drive-root names without a file name", () => ReuseSelectionAsync(null, ["Local Disk (C:)"],
            capturedItems: [("Local Disk (C:)", @"C:\", true)]));
    }

    private static async Task ReuseSelectionAsync(string[]? requested, string[] expected, bool expectedReused = true, bool viewAvailable = true,
        (string Name, string Path, bool IsFileSystem)[]? capturedItems = null)
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(async () =>
        {
            using var lifetime = new CancellationTokenSource();
            using var openLock = new SemaphoreSlim(1);
            using var pathComparer = new ShellPathComparer();
            using var fixture = new ExplorerTabActivationTests.ActivationWindow();
            var watcher = ExplorerTabActivationTests.CreateSelectionWatcher(lifetime);
            SetField(watcher, "_toOpenWindowsLock", openLock);
            SetField(watcher, "_windowEntryDictLock", new object());
            SetField(watcher, "_staTaskScheduler", scheduler);
            SetField(watcher, "_shellPathComparer", pathComparer);
            SetField(watcher, "_preExistingExplorerWindowsProtected", true);
            SetField(watcher, "_reuseTabs", true);
            SetField(watcher, "_isForcingTabs", true);
            SetField(watcher, "_hookLifetime", lifetime);
            SetField(watcher, "_mergeSourceHWnds", Activator.CreateInstance(typeof(ExplorerWatcher)
                .GetField("_mergeSourceHWnds", BindingFlags.Instance | BindingFlags.NonPublic)!.FieldType)!);
            SetField(watcher, "_closingMergeSourceHWnds", new System.Collections.Concurrent.ConcurrentDictionary<nint, MergeOperation>());

            var selected = new List<string> { "previous.txt" };
            var view = viewAvailable ? CreateFolderView(selected, capturedItems) : null;
            var dictionaryField = typeof(ExplorerWatcher).GetField("_windowEntryDict", BindingFlags.Instance | BindingFlags.NonPublic)!;
            var dictionary = Activator.CreateInstance(dictionaryField.FieldType)!;
            var browserType = dictionaryField.FieldType.GetGenericArguments()[0];
            var browser = ShellDispatchStub.Create(browserType, (method, arguments) => method switch
            {
                "get_HWND" => (long)fixture.Handle,
                "get_LocationURL" => "file:///C:/WinTab-reuse",
                "get_Document" => view,
                _ => throw new InvalidOperationException("Unexpected browser call: " + method)
            });
            // Exercise the handoff from the source selection to the reused tab. Display names may
            // omit extensions, but Explorer's folder parser still needs the complete file name.
            if (capturedItems != null)
                requested = (string[]?)typeof(ExplorerWatcher).GetMethod("GetSelectedItems", BindingFlags.Static | BindingFlags.NonPublic)!
                    .Invoke(null, [browser]);
            var location = Helper.NormalizeLocation("C:/WinTab-reuse");
            var info = new WindowInfo
            {
                Identity = WindowIdentity.Capture(fixture.Handle),
                TabIdentity = WindowIdentity.Capture(fixture.FirstTab),
                HookedTopLevelHWnd = fixture.Handle,
                EventsHooked = true,
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

    private static object CreateFolderView(List<string> selected, (string Name, string Path, bool IsFileSystem)[]? capturedItems)
    {
        var viewType = typeof(ExplorerWatcher).Assembly.GetType("Shell32.ShellFolderView", throwOnError: true)!;
        var folderType = viewType.GetInterfaces().Append(viewType).SelectMany(type => type.GetProperties())
            .First(property => property.Name == "Folder").PropertyType;
        var itemType = folderType.GetInterfaces().Append(folderType).SelectMany(type => type.GetMethods())
            .First(method => method.Name == "ParseName").ReturnType;
        var selectionType = viewType.GetInterfaces().Append(viewType).SelectMany(type => type.GetMethods())
            .First(method => method.Name == "SelectedItems").ReturnType;
        var sourceItems = (capturedItems ?? []).Select(source => ShellDispatchStub.Create(itemType, (method, arguments) => method switch
        {
            "get_Name" => source.Name,
            "get_Path" => source.Path,
            "get_IsFileSystem" => source.IsFileSystem,
            _ => throw new InvalidOperationException("Unexpected source item call: " + method)
        })).ToArray();
        var selection = ShellDispatchStub.Create(selectionType, (method, arguments) => method switch
        {
            "get_Count" => sourceItems.Length,
            "Item" => sourceItems[Convert.ToInt32(arguments[0])],
            _ => throw new InvalidOperationException("Unexpected selection call: " + method)
        });
        // Model real ParseName behavior: a display label is not an alias for a file with an extension.
        var availableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "previous.txt", "target.txt", "first.txt", "second.txt", "测试应用.lnk", "report.txt", "Control Panel", "Local Disk (C:)" };
        var parsedNames = new Dictionary<object, string>();
        var folder = ShellDispatchStub.Create(folderType, (method, arguments) =>
        {
            if (method != "ParseName")
                throw new InvalidOperationException("Unexpected folder call: " + method);
            var name = (string)arguments[0]!;
            if (!availableNames.Contains(name))
                return null;
            var item = ShellDispatchStub.Create(itemType, (itemMethod, itemArguments) =>
                itemMethod == "get_Name" ? name : throw new InvalidOperationException("Unexpected item call: " + itemMethod));
            parsedNames.Add(item, name);
            return item;
        });
        return ShellDispatchStub.Create(viewType, (method, arguments) =>
        {
            if (method == "SelectedItems")
                return selection;
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

    protected override object? Invoke(MethodInfo? targetMethod, object?[]? arguments)
    {
        var result = _invoke(targetMethod!.Name, arguments ?? []);
        if (result != null && targetMethod != null && targetMethod.ReturnType != typeof(void) && !targetMethod.ReturnType.IsInstanceOfType(result))
        {
            if (targetMethod.ReturnType == typeof(long) && result is int intVal)
                return (long)intVal;
            return Convert.ChangeType(result, targetMethod.ReturnType);
        }
        return result;
    }
}
