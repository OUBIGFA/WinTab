using System;
using System.Reflection;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public sealed class OpenTargetRoutingTests
{
    [Fact]
    public void OpenTargetClassifier_ShouldClassifyRecycleBinAliasesAsNativeShellNamespace()
    {
        object result = InvokeClassify("shell:RecycleBinFolder");
        string kind = ReadKind(result);

        kind.Should().Be("ShellNamespace");
    }

    [Fact]
    public void OpenTargetClassifier_ShouldClassifyPhysicalFolderAsPhysicalFileSystem()
    {
        object result = InvokeClassify(@"C:\Windows");
        string kind = ReadKind(result);

        kind.Should().Be("PhysicalFileSystem");
    }

    [Fact]
    public void OpenTargetClassifier_ShouldKeepNonRecycleShellNamespaceOutOfNativeOnlyBucket()
    {
        object result = InvokeClassify("shell::Downloads");
        string kind = ReadKind(result);

        kind.Should().Be("ShellNamespace");
    }

    private static object InvokeClassify(string target)
    {
        Type? classifierType = Type.GetType(
            "WinTab.Platform.Win32.OpenTargetClassifier, WinTab.Platform.Win32",
            throwOnError: false);

        classifierType.Should().NotBeNull("the refactor must provide a shared classifier used by App and ShellBridge");

        MethodInfo? classify = classifierType!.GetMethod(
            "Classify",
            BindingFlags.Public | BindingFlags.Static,
            binder: null,
            types: [typeof(string)],
            modifiers: null);

        classify.Should().NotBeNull("target routing must expose a stable classification entry point");

        object? result = classify!.Invoke(null, [target]);
        result.Should().NotBeNull();
        return result!;
    }

    private static string ReadKind(object result)
    {
        PropertyInfo? kindProperty = result.GetType().GetProperty("Kind", BindingFlags.Public | BindingFlags.Instance);
        kindProperty.Should().NotBeNull("classification result must expose the routed target kind");

        object? kind = kindProperty!.GetValue(result);
        kind.Should().NotBeNull();
        return kind!.ToString() ?? string.Empty;
    }
}
