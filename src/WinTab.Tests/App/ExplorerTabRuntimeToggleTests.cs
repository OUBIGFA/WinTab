using System.Reflection;
using FluentAssertions;
using WinTab.App.ExplorerTabUtilityPort;
using Xunit;

namespace WinTab.Tests.App;

public sealed class ExplorerTabRuntimeToggleTests
{
    [Fact]
    public void ExplorerTabHookService_ShouldExposeRuntimeAutoConvertSwitch()
    {
        MethodInfo? method = typeof(ExplorerTabHookService).GetMethod(
            "SetAutoConvertEnabled",
            BindingFlags.Public | BindingFlags.Instance);

        method.Should().NotBeNull(
            "auto-convert toggle changes should be applied immediately without restart");
    }

    [Fact]
    public void BehaviorViewModel_ShouldDependOnRuntimeAutoConvertController()
    {
        ConstructorInfo? ctor = typeof(WinTab.App.ViewModels.BehaviorViewModel)
            .GetConstructors(BindingFlags.Public | BindingFlags.Instance)
            .OrderByDescending(c => c.GetParameters().Length)
            .FirstOrDefault();

        ctor.Should().NotBeNull();

        string[] parameterTypeNames = ctor!.GetParameters()
            .Select(p => p.ParameterType.Name)
            .ToArray();

        parameterTypeNames.Should().Contain("IExplorerAutoConvertController",
            "view-model setting changes should push runtime state into hook service");
    }
}
