using System.IO;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public class PageLocalizationCoverageTests
{
    private static readonly Regex DynamicResourceRegex = new("\\{DynamicResource\\s+([^}\\s]+)\\}", RegexOptions.Compiled);

    [Fact]
    public void GeneralAndBehaviorPages_DynamicResourceKeys_ShouldExistInAllLocalizationFiles()
    {
        string srcRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "src"));

        string generalPagePath = Path.Combine(srcRoot, "WinTab.App", "Views", "Pages", "GeneralPage.xaml");
        string behaviorPagePath = Path.Combine(srcRoot, "WinTab.App", "Views", "Pages", "BehaviorPage.xaml");
        string zhPath = Path.Combine(srcRoot, "WinTab.UI", "Localization", "Strings.zh-CN.xaml");
        string enPath = Path.Combine(srcRoot, "WinTab.UI", "Localization", "Strings.en-US.xaml");

        var resourceKeys = ExtractDynamicResourceKeys(generalPagePath)
            .Union(ExtractDynamicResourceKeys(behaviorPagePath), StringComparer.Ordinal)
            .ToList();

        Dictionary<string, string> zhResources = ReadStringResources(zhPath);
        Dictionary<string, string> enResources = ReadStringResources(enPath);

        foreach (string key in resourceKeys)
        {
            zhResources.ContainsKey(key).Should().BeTrue($"zh-CN localization must contain key '{key}'");
            enResources.ContainsKey(key).Should().BeTrue($"en-US localization must contain key '{key}'");

            zhResources[key].Should().NotBeNullOrWhiteSpace($"zh-CN localization value for '{key}' must not be empty");
            enResources[key].Should().NotBeNullOrWhiteSpace($"en-US localization value for '{key}' must not be empty");
        }
    }

    [Fact]
    public void AppXaml_TextBlockStyles_ShouldBeBasedOnDefaultTextBlockStyle()
    {
        string appXamlPath = Path.GetFullPath(
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "src", "WinTab.App", "App.xaml"));

        XDocument appXaml = XDocument.Load(appXamlPath);
        var styleElements = appXaml.Descendants().Where(e => e.Name.LocalName == "Style").ToList();

        string[] textStyles = ["PageSectionTitleStyle", "PageGroupLabelStyle", "SettingTitleStyle", "SettingDescriptionStyle"];

        foreach (string styleKey in textStyles)
        {
            XElement? style = styleElements.FirstOrDefault(e =>
                string.Equals((string?)e.Attribute(XName.Get("Key", "http://schemas.microsoft.com/winfx/2006/xaml")), styleKey, StringComparison.Ordinal));

            style.Should().NotBeNull($"App.xaml should define style '{styleKey}'");
            style!.Attribute("BasedOn").Should().NotBeNull($"Text style '{styleKey}' must preserve theme defaults via BasedOn");
        }
    }

    private static Dictionary<string, string> ReadStringResources(string path)
    {
        XDocument doc = XDocument.Load(path);
        XNamespace x = "http://schemas.microsoft.com/winfx/2006/xaml";

        return doc
            .Descendants()
            .Where(e => e.Name.LocalName == "String")
            .Where(e => e.Attribute(x + "Key") is not null)
            .ToDictionary(
                e => e.Attribute(x + "Key")!.Value,
                e => e.Value,
                StringComparer.Ordinal);
    }

    private static IEnumerable<string> ExtractDynamicResourceKeys(string pagePath)
    {
        XDocument page = XDocument.Load(pagePath);

        foreach (XElement element in page.Descendants())
        {
            foreach (XAttribute attr in element.Attributes())
            {
                if (!string.Equals(attr.Name.LocalName, "Text", StringComparison.Ordinal) &&
                    !string.Equals(attr.Name.LocalName, "Content", StringComparison.Ordinal))
                {
                    continue;
                }

                Match match = DynamicResourceRegex.Match(attr.Value);
                if (match.Success)
                {
                    yield return match.Groups[1].Value;
                }
            }
        }
    }
}
