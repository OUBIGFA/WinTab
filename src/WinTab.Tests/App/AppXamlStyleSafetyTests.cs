using System.IO;
using System.Xml.Linq;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public class AppXamlStyleSafetyTests
{
    [Fact]
    public void AppXaml_ShouldNotDefineGlobalImplicitStylesForWpfUiControls()
    {
        string appXamlPath = GetProjectFilePath("WinTab.App", "App.xaml");

        File.Exists(appXamlPath).Should().BeTrue("App.xaml must exist for style safety validation");

        XDocument appXaml = XDocument.Load(appXamlPath);
        var styleElements = appXaml.Descendants().Where(e => e.Name.LocalName == "Style").ToList();

        string[] highRiskTargets = ["ui:CardControl", "ui:Button"];

        foreach (string targetType in highRiskTargets)
        {
            var implicitStyles = styleElements
                .Where(style => string.Equals((string?)style.Attribute("TargetType"), targetType, StringComparison.Ordinal))
                .Where(style => style.Attribute(XName.Get("Key", "http://schemas.microsoft.com/winfx/2006/xaml")) is null)
                .ToList();

            implicitStyles.Should().BeEmpty(
                $"global implicit style for {targetType} can override WPF-UI control templates and hide header text");
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
