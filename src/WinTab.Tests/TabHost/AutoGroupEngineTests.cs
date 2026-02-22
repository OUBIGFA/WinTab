using WinTab.Core.Enums;
using WinTab.Core.Models;
using WinTab.TabHost;
using Xunit;

namespace WinTab.Tests.TabHost;

public sealed class AutoGroupEngineTests
{
    [Fact]
    public void IsMatch_ProcessName_IsCaseInsensitive()
    {
        var rule = new AutoGroupRule
        {
            MatchType = AutoGroupMatchType.ProcessName,
            MatchValue = "notepad.exe"
        };

        var window = new WindowInfo(
            Handle: (IntPtr)123,
            Title: "Untitled - Notepad",
            ProcessName: "Notepad.EXE",
            ProcessId: 42,
            ClassName: "Notepad",
            IsVisible: true);

        bool matched = AutoGroupEngine.IsMatch(rule, window);

        Assert.True(matched);
    }

    [Fact]
    public void IsMatch_InvalidRegex_ReturnsFalse()
    {
        var rule = new AutoGroupRule
        {
            MatchType = AutoGroupMatchType.WindowTitleRegex,
            MatchValue = "[unterminated"
        };

        var window = new WindowInfo(
            Handle: (IntPtr)1,
            Title: "Chrome - Chat",
            ProcessName: "chrome.exe",
            ProcessId: 100,
            ClassName: "Chrome_WidgetWin_1",
            IsVisible: true);

        bool matched = AutoGroupEngine.IsMatch(rule, window);

        Assert.False(matched);
    }

    [Fact]
    public void IsMatch_ProcessPath_UsesContains()
    {
        var rule = new AutoGroupRule
        {
            MatchType = AutoGroupMatchType.ProcessPath,
            MatchValue = "\\Google\\Chrome\\"
        };

        var window = new WindowInfo(
            Handle: (IntPtr)88,
            Title: "Docs",
            ProcessName: "chrome.exe",
            ProcessId: 88,
            ClassName: "Chrome_WidgetWin_1",
            IsVisible: true,
            ProcessPath: @"C:\Program Files\Google\Chrome\Application\chrome.exe");

        bool matched = AutoGroupEngine.IsMatch(rule, window);

        Assert.True(matched);
    }
}
