using System.IO;
using System.Xml.Linq;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public sealed class SettingRowAlignmentRegressionTests
{
    [Theory]
    [InlineData("GeneralPage.xaml", 5)]
    [InlineData("BehaviorPage.xaml", 4)]
    [InlineData("AboutPage.xaml", 2)]
    [InlineData("UninstallPage.xaml", 1)]
    public void SettingsPages_ShouldUseSharedActionColumnForConsistentRowAlignment(string pageName, int expectedRowCount)
    {
        string pagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", pageName);
        XDocument page = XDocument.Load(pagePath);

        XElement contentStack = page
            .Descendants()
            .First(e => e.Name.LocalName == "StackPanel" && e.Attribute("MaxWidth") is not null);

        contentStack.Attribute("Grid.IsSharedSizeScope")?.Value.Should().Be("True",
            "the polished settings layout should synchronize right-side control columns within each page");

        var actionColumns = page
            .Descendants()
            .Where(e => e.Name.LocalName == "ColumnDefinition"
                && (string?)e.Attribute("SharedSizeGroup") == "SettingActionColumn")
            .ToList();

        actionColumns.Should().HaveCount(expectedRowCount,
            "every settings row should share the same accessory column width");
    }

    [Fact]
    public void DesignPrimitives_ShouldProvideSharedActionHostStyle()
    {
        string primitivesPath = GetProjectFilePath("WinTab.UI", "Themes", "DesignPrimitives.xaml");
        XDocument primitives = XDocument.Load(primitivesPath);

        XElement style = primitives
            .Descendants()
            .First(e => e.Name.LocalName == "Style"
                && (string?)e.Attribute(XName.Get("Key", "http://schemas.microsoft.com/winfx/2006/xaml")) == "SettingRowActionHostStyle");

        var setters = style
            .Descendants()
            .Where(e => e.Name.LocalName == "Setter")
            .ToDictionary(e => (string)e.Attribute("Property")!, e => (string?)e.Attribute("Value"));

        setters["HorizontalAlignment"].Should().Be("Stretch");
        setters["VerticalAlignment"].Should().Be("Center");
        setters["Margin"].Should().Be("24,0,0,0");
        setters["MinWidth"].Should().Be("220");
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
