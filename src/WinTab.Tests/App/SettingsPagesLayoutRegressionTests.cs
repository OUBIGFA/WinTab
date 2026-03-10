using System.IO;
using System.Xml.Linq;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public sealed class SettingsPagesLayoutRegressionTests
{
    [Fact]
    public void GeneralAndBehaviorPages_ShouldUseDedicatedSettingRowsInsteadOfCardControlHeaders()
    {
        string generalPagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", "GeneralPage.xaml");
        string behaviorPagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", "BehaviorPage.xaml");

        XDocument generalPage = XDocument.Load(generalPagePath);
        XDocument behaviorPage = XDocument.Load(behaviorPagePath);

        generalPage.Descendants().Count(e => e.Name.LocalName == "CardControl").Should().Be(0,
            "the startup page labels should not depend on CardControl header rendering");
        behaviorPage.Descendants().Count(e => e.Name.LocalName == "CardControl").Should().Be(0,
            "the tab rules page labels should not depend on CardControl header rendering");
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
