using System.IO;
using System.Xml.Linq;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public sealed class MainWindowShellHeaderRegressionTests
{
    [Fact]
    public void MainWindow_ShouldNotRenderLegacyStatusPillsInShellHeader()
    {
        string pagePath = GetProjectFilePath("WinTab.App", "Views", "MainWindow.xaml");
        XDocument page = XDocument.Load(pagePath);

        var textResources = page
            .Descendants()
            .Where(e => e.Name.LocalName == "TextBlock")
            .Select(e => (string?)e.Attribute("Text"))
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .ToList();

        textResources.Should().NotContain("{DynamicResource Shell_StatusPrimary}",
            "the shell header should no longer show redundant platform status pills");
        textResources.Should().NotContain("{DynamicResource Shell_StatusSecondary}",
            "the shell header should no longer show redundant theme/language status pills");
    }

    private static string GetProjectFilePath(string projectFolder, params string[] parts)
    {
        string[] allParts = ["src", projectFolder, .. parts];
        return TestRepoPaths.GetFile(allParts);
    }
}
