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
        string current = AppContext.BaseDirectory;

        for (int i = 0; i < 8; i++)
        {
            string candidate = Path.Combine(current, "src", projectFolder);
            if (Directory.Exists(candidate))
            {
                return Path.Combine(candidate, Path.Combine(parts));
            }

            string? parent = Directory.GetParent(current)?.FullName;
            if (string.IsNullOrWhiteSpace(parent))
            {
                break;
            }

            current = parent;
        }

        return Path.Combine(AppContext.BaseDirectory, Path.Combine(parts));
    }
}
