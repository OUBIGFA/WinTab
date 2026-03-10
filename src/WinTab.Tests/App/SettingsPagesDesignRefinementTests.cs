using System.IO;
using System.Xml.Linq;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public class SettingsPagesDesignRefinementTests
{
    [Fact]
    public void GeneralAndBehaviorPages_ShouldKeepRefinedLayoutMarkers()
    {
        string generalPagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", "GeneralPage.xaml");
        string behaviorPagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", "BehaviorPage.xaml");

        XDocument generalPage = XDocument.Load(generalPagePath);
        XDocument behaviorPage = XDocument.Load(behaviorPagePath);

        string? generalMaxWidth = generalPage
            .Descendants()
            .First(e => e.Name.LocalName == "StackPanel" && e.Attribute("MaxWidth") is not null)
            .Attribute("MaxWidth")?.Value;

        string? behaviorMaxWidth = behaviorPage
            .Descendants()
            .First(e => e.Name.LocalName == "StackPanel" && e.Attribute("MaxWidth") is not null)
            .Attribute("MaxWidth")?.Value;

        generalMaxWidth.Should().Be("880", "General page should use the wider polished settings rhythm");
        behaviorMaxWidth.Should().Be("880", "Behavior page should use the wider polished settings rhythm");

        var generalRows = generalPage
            .Descendants()
            .Where(e => e.Name.LocalName == "Border" && (string?)e.Attribute("Style") == "{StaticResource SettingRowBorderStyle}")
            .ToList();
        var behaviorRows = behaviorPage
            .Descendants()
            .Where(e => e.Name.LocalName == "Border" && (string?)e.Attribute("Style") == "{StaticResource SettingRowBorderStyle}")
            .ToList();

        generalRows.Should().HaveCount(5, "General page should render five dedicated setting rows");
        behaviorRows.Should().HaveCount(4, "Behavior page should render four dedicated setting rows");

        foreach (XElement row in generalRows)
        {
            row.Attribute("Style")?.Value.Should().Be("{StaticResource SettingRowBorderStyle}", "General rows should share the dedicated setting row style");
        }

        foreach (XElement row in behaviorRows)
        {
            row.Attribute("Style")?.Value.Should().Be("{StaticResource SettingRowBorderStyle}", "Behavior rows should share the dedicated setting row style");
        }
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
