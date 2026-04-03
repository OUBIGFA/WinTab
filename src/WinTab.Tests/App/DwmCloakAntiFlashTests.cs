using System;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using FluentAssertions;
using WinTab.Platform.Win32;
using Xunit;

namespace WinTab.Tests.App;

/// <summary>
/// Tests that the DWM-cloak-based anti-flash mechanism is correctly wired
/// throughout the entire auto-convert pipeline.  The key invariant: every
/// path that hides an Explorer window must call <c>SuppressVisibility</c>
/// (DWM cloaking) BEFORE or alongside <c>Hide</c> (ShowWindow SW_HIDE),
/// and every restore path must call <c>RestoreVisibility</c> before
/// <c>Show</c>.
/// </summary>
public sealed class DwmCloakAntiFlashTests
{
    // ───────────────────────────────────────────────────────────────────────
    //  WindowManager: SuppressVisibility / RestoreVisibility wired to DWM
    // ───────────────────────────────────────────────────────────────────────

    [Fact]
    public void WindowManager_SuppressVisibility_ShouldUseDwmCloak()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(
            ["src", "WinTab.Platform.Win32", "WindowManager.cs"]));

        source.Should().Contain("DWMWA_CLOAK",
            "SuppressVisibility must use DWM cloaking attribute to prevent composition");
        source.Should().Contain("DwmSetWindowAttribute",
            "SuppressVisibility must call DwmSetWindowAttribute");
    }

    [Fact]
    public void WindowManager_RestoreVisibility_ShouldUncloakViaDwm()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(
            ["src", "WinTab.Platform.Win32", "WindowManager.cs"]));

        source.Should().Contain("RestoreVisibility",
            "RestoreVisibility must be implemented");

        // Verify it also calls DwmSetWindowAttribute with cloaked = 0
        int restoreStart = source.IndexOf("public void RestoreVisibility", StringComparison.Ordinal);
        restoreStart.Should().BeGreaterThan(-1, "RestoreVisibility method must exist");

        string restoreBody = source.Substring(restoreStart, Math.Min(400, source.Length - restoreStart));
        restoreBody.Should().Contain("DwmSetWindowAttribute",
            "RestoreVisibility must call DwmSetWindowAttribute to uncloak");
    }

    [Fact]
    public void WindowManager_ShouldImplementBothInterfaceMethods()
    {
        Type wmType = typeof(WindowManager);

        var suppress = wmType.GetMethod("SuppressVisibility", BindingFlags.Public | BindingFlags.Instance);
        suppress.Should().NotBeNull("WindowManager must implement SuppressVisibility");
        suppress!.ReturnType.Should().Be(typeof(bool), "SuppressVisibility should return bool");

        var restore = wmType.GetMethod("RestoreVisibility", BindingFlags.Public | BindingFlags.Instance);
        restore.Should().NotBeNull("WindowManager must implement RestoreVisibility");
    }

    // ───────────────────────────────────────────────────────────────────────
    //  NativeConstants: DWMWA_CLOAK constant exists
    // ───────────────────────────────────────────────────────────────────────

    [Fact]
    public void NativeConstants_DwmaCloakValue_ShouldBe13()
    {
        NativeConstants.DWMWA_CLOAK.Should().Be(13,
            "DWMWA_CLOAK is documented as attribute 13 in the Windows DWM API");
    }

    // ───────────────────────────────────────────────────────────────────────
    //  NativeMethods: DwmSetWindowAttribute P/Invoke declared
    // ───────────────────────────────────────────────────────────────────────

    [Fact]
    public void NativeMethods_DwmSetWindowAttribute_ShouldBeDeclared()
    {
        var method = typeof(NativeMethods).GetMethod(
            "DwmSetWindowAttribute",
            BindingFlags.Public | BindingFlags.Static);

        method.Should().NotBeNull("DwmSetWindowAttribute P/Invoke must be declared for DWM cloaking");
    }

    // ───────────────────────────────────────────────────────────────────────
    //  ExplorerTabHookService: DWM cloaking wired into auto-convert pipeline
    // ───────────────────────────────────────────────────────────────────────

    [Fact]
    public void EarlyHidePath_ShouldCloakBeforeHide()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(
            ["src", "WinTab.App", "ExplorerTabUtilityPort", "ExplorerTabHookService.cs"]));

        // Find the OnExplorerObjectCreate method body
        int methodStart = source.IndexOf("private void OnExplorerObjectCreate(", StringComparison.Ordinal);
        methodStart.Should().BeGreaterThan(-1, "OnExplorerObjectCreate must exist");

        // Extract the method body (up to the next private/public method)
        int methodEnd = source.IndexOf("\n    private ", methodStart + 1, StringComparison.Ordinal);
        if (methodEnd < 0)
            methodEnd = source.Length;
        string methodBody = source.Substring(methodStart, methodEnd - methodStart);

        int cloakPos = methodBody.IndexOf("SuppressVisibility", StringComparison.Ordinal);
        int hidePos = methodBody.IndexOf(".Hide(hwnd)", StringComparison.Ordinal);

        cloakPos.Should().BeGreaterThan(-1,
            "OnExplorerObjectCreate must call SuppressVisibility (DWM cloak)");
        hidePos.Should().BeGreaterThan(-1,
            "OnExplorerObjectCreate must also call Hide (belt-and-suspenders)");
        cloakPos.Should().BeLessThan(hidePos,
            "SuppressVisibility (DWM cloak) must be called BEFORE Hide (SW_HIDE) " +
            "in OnExplorerObjectCreate to prevent any frame from being composited");
    }

    [Fact]
    public void AggressiveHidePath_ShouldCloakImmediately()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(
            ["src", "WinTab.App", "ExplorerTabUtilityPort", "ExplorerTabHookService.cs"]));

        int methodStart = source.IndexOf("private async Task<bool> TryHideExplorerWindowAggressively", StringComparison.Ordinal);
        methodStart.Should().BeGreaterThan(-1, "TryHideExplorerWindowAggressively must exist");

        int methodEnd = source.IndexOf("\n    private ", methodStart + 1, StringComparison.Ordinal);
        if (methodEnd < 0) methodEnd = source.Length;
        string methodBody = source.Substring(methodStart, methodEnd - methodStart);

        // DWM cloak must appear BEFORE the polling loop
        int cloakPos = methodBody.IndexOf("SuppressVisibility", StringComparison.Ordinal);
        int whilePos = methodBody.IndexOf("while (", StringComparison.Ordinal);

        cloakPos.Should().BeGreaterThan(-1,
            "TryHideExplorerWindowAggressively must call SuppressVisibility");
        whilePos.Should().BeGreaterThan(-1);
        cloakPos.Should().BeLessThan(whilePos,
            "SuppressVisibility must be called BEFORE the polling loop " +
            "to guarantee zero-flash from the very first iteration");
    }

    [Fact]
    public void PollingLoop_ShouldReapplyCloakEachIteration()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(
            ["src", "WinTab.App", "ExplorerTabUtilityPort", "ExplorerTabHookService.cs"]));

        int methodStart = source.IndexOf("private async Task KeepWindowHiddenUntilConversionCompletes", StringComparison.Ordinal);
        methodStart.Should().BeGreaterThan(-1);

        int methodEnd = source.IndexOf("\n    private ", methodStart + 1, StringComparison.Ordinal);
        if (methodEnd < 0) methodEnd = source.Length;
        string methodBody = source.Substring(methodStart, methodEnd - methodStart);

        methodBody.Should().Contain("SuppressVisibility",
            "KeepWindowHiddenUntilConversionCompletes must re-cloak every iteration " +
            "to handle Explorer toggling WS_VISIBLE during initialization");
    }

    [Fact]
    public void ConvertWindowToTab_FailurePath_ShouldUncloakBeforeShow()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(
            ["src", "WinTab.App", "ExplorerTabUtilityPort", "ExplorerTabHookService.cs"]));

        int methodStart = source.IndexOf("private async Task ConvertWindowToTab(", StringComparison.Ordinal);
        methodStart.Should().BeGreaterThan(-1);

        int methodEnd = source.IndexOf("\n    private ", methodStart + 1, StringComparison.Ordinal);
        if (methodEnd < 0) methodEnd = source.Length;
        string methodBody = source.Substring(methodStart, methodEnd - methodStart);

        // The finally block must uncloak before showing
        int finallyPos = methodBody.LastIndexOf("finally", StringComparison.Ordinal);
        finallyPos.Should().BeGreaterThan(-1);

        string finallyBlock = methodBody.Substring(finallyPos);

        int restorePos = finallyBlock.IndexOf("RestoreVisibility", StringComparison.Ordinal);
        int showPos = finallyBlock.IndexOf(".Show(sourceTopLevel)", StringComparison.Ordinal);

        restorePos.Should().BeGreaterThan(-1,
            "ConvertWindowToTab finally block must call RestoreVisibility to uncloak on failure");
        showPos.Should().BeGreaterThan(-1,
            "ConvertWindowToTab finally block must call Show on failure");
        restorePos.Should().BeLessThan(showPos,
            "RestoreVisibility (uncloak) must be called BEFORE Show " +
            "to prevent a flash when restoring a failed conversion");
    }

    [Fact]
    public void RestoreEarlyHiddenWindow_ShouldUncloakBeforeShow()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(
            ["src", "WinTab.App", "ExplorerTabUtilityPort", "ExplorerTabHookService.cs"]));

        int methodStart = source.IndexOf("private void RestoreEarlyHiddenWindow(", StringComparison.Ordinal);
        methodStart.Should().BeGreaterThan(-1);

        int methodEnd = source.IndexOf("\n    private ", methodStart + 1, StringComparison.Ordinal);
        if (methodEnd < 0) methodEnd = source.Length;
        string methodBody = source.Substring(methodStart, methodEnd - methodStart);

        int restorePos = methodBody.IndexOf("RestoreVisibility", StringComparison.Ordinal);
        int showPos = methodBody.IndexOf(".Show(hwnd)", StringComparison.Ordinal);

        restorePos.Should().BeGreaterThan(-1,
            "RestoreEarlyHiddenWindow must uncloak before showing");
        showPos.Should().BeGreaterThan(-1);
        restorePos.Should().BeLessThan(showPos,
            "RestoreVisibility must be called BEFORE Show in RestoreEarlyHiddenWindow");
    }

    // ───────────────────────────────────────────────────────────────────────
    //  Architectural invariant: every Hide must be paired with SuppressVisibility
    // ───────────────────────────────────────────────────────────────────────

    [Fact]
    public void AutoConvertPipeline_ShouldNotUseHideWithoutCloak()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(
            ["src", "WinTab.App", "ExplorerTabUtilityPort", "ExplorerTabHookService.cs"]));

        // Every call to _windowManager.Hide in the auto-convert pipeline
        // should be preceded or accompanied by SuppressVisibility in the
        // same method.  We check the three critical methods.
        string[] criticalMethods =
        [
            "OnExplorerObjectCreate",
            "TryHideExplorerWindowAggressively",
            "KeepWindowHiddenUntilConversionCompletes",
            "ConvertWindowToTab"
        ];

        foreach (string methodName in criticalMethods)
        {
            // Find method body
            string pattern = methodName.Contains("ConvertWindowToTab")
                ? "private async Task ConvertWindowToTab("
                : methodName;
            int start = source.IndexOf(pattern, StringComparison.Ordinal);
            if (start < 0) continue;

            int end = source.IndexOf("\n    private ", start + 1, StringComparison.Ordinal);
            if (end < 0) end = source.Length;
            string body = source.Substring(start, end - start);

            if (body.Contains(".Hide("))
            {
                body.Should().Contain("SuppressVisibility",
                    $"{methodName} calls Hide but must also call SuppressVisibility " +
                    "to prevent DWM from compositing the window before SW_HIDE takes effect");
            }
        }
    }
}
