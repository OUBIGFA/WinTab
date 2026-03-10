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

        var generalCards = generalPage.Descendants().Where(e => e.Name.LocalName == "CardControl").ToList();
        var behaviorCards = behaviorPage.Descendants().Where(e => e.Name.LocalName == "CardControl").ToList();

        generalCards.Should().NotBeEmpty();
        behaviorCards.Should().NotBeEmpty();

        foreach (XElement card in generalCards)
        {
            card.Attribute("Style")?.Value.Should().Be("{StaticResource SettingsCardStyle}", "General cards should share the polished card style token");
        }

        foreach (XElement card in behaviorCards)
        {
            card.Attribute("Style")?.Value.Should().Be("{StaticResource SettingsCardStyle}", "Behavior cards should share the polished card style token");
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
