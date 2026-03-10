using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;
using FluentAssertions;
using WinTab.App.Views.Pages;
using WinTab.Core.Enums;
using WinTab.UI.Localization;
using Xunit;

namespace WinTab.Tests.App;

public class AboutPageRegressionTests
{
    [Fact]
    public void AboutPage_Runtime_ShouldRenderMultipleNonEmptyTextBlocks()
    {
        int nonEmptyTextCount = RunInSta(() =>
        {
            EnsureApplication();
            LocalizationManager.ApplyLanguage(Language.Chinese);

            var page = new AboutPage(null!);
            return RenderAndCountNonEmptyTextBlocks(page);
        });

        nonEmptyTextCount.Should().BeGreaterThan(3, "About page should display app name, description and section texts even with design-time view model");
    }

    [Fact]
    public void AboutPage_Design_ShouldAlignWithPolishedProjectAndDiagnosticsComposition()
    {
        string pagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", "AboutPage.xaml");
        XDocument page = XDocument.Load(pagePath);

        string? maxWidth = page
            .Descendants()
            .First(e => e.Name.LocalName == "StackPanel" && e.Attribute("MaxWidth") is not null)
            .Attribute("MaxWidth")?.Value;

        maxWidth.Should().Be("860", "polished About layout should use a roomier centered content width");

        string? footerAlignment = page
            .Descendants()
            .First(e => e.Name.LocalName == "StackPanel" && (string?)e.Attribute("Grid.Row") == "1")
            .Attribute("HorizontalAlignment")?.Value;

        footerAlignment.Should().Be("Center", "About footer should stay centered in the polished layout");
    }

    [Fact]
    public void AboutPage_ShouldPromoteLeadCopyAndProjectSection()
    {
        string pagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", "AboutPage.xaml");
        XDocument page = XDocument.Load(pagePath);

        XElement lead = page
            .Descendants()
            .First(e => e.Name.LocalName == "TextBlock" && (string?)e.Attribute("Text") == "{DynamicResource About_Lead}");

        lead.Attribute("Style")?.Value.Should().Be("{StaticResource PageHeaderDescriptionStyle}", "About page should promote lead copy as a readable wrapped introduction");

        bool hasProjectSection = page
            .Descendants()
            .Any(e => e.Name.LocalName == "TextBlock" && (string?)e.Attribute("Text") == "{DynamicResource About_SectionProject}");

        hasProjectSection.Should().BeTrue("About page should expose a dedicated project section in the polished layout");
    }

    private static int RenderAndCountNonEmptyTextBlocks(Page page)
    {
        var host = new Window
        {
            Width = 1000,
            Height = 760,
            Content = page,
            ShowInTaskbar = false,
            WindowStyle = WindowStyle.None,
            AllowsTransparency = true,
            Opacity = 0
        };

        host.Show();
        host.UpdateLayout();
        int count = CountNonEmptyTextBlocks(page);
        host.Hide();
        host.Close();

        return count;
    }

    private static int CountNonEmptyTextBlocks(DependencyObject root)
    {
        int count = 0;

        foreach (TextBlock textBlock in FindLogicalChildren<TextBlock>(root))
        {
            if (!string.IsNullOrWhiteSpace(textBlock.Text))
            {
                count++;
            }
        }

        foreach (Wpf.Ui.Controls.CardControl card in FindLogicalChildren<Wpf.Ui.Controls.CardControl>(root))
        {
            count += CountNonEmptyTextBlocksFromObject(card.Header);
        }

        return count;
    }

    private static int CountNonEmptyTextBlocksFromObject(object? value)
    {
        if (value is null)
        {
            return 0;
        }

        if (value is TextBlock textBlock)
        {
            return string.IsNullOrWhiteSpace(textBlock.Text) ? 0 : 1;
        }

        if (value is DependencyObject dependencyObject)
        {
            int count = 0;
            foreach (TextBlock nestedTextBlock in FindLogicalChildren<TextBlock>(dependencyObject))
            {
                if (!string.IsNullOrWhiteSpace(nestedTextBlock.Text))
                {
                    count++;
                }
            }

            return count;
        }

        return 0;
    }

    private static IEnumerable<T> FindLogicalChildren<T>(DependencyObject root)
        where T : DependencyObject
    {
        if (root is null)
        {
            yield break;
        }

        foreach (object child in LogicalTreeHelper.GetChildren(root))
        {
            if (child is not DependencyObject childObject)
            {
                continue;
            }

            if (childObject is T target)
            {
                yield return target;
            }

            foreach (T nested in FindLogicalChildren<T>(childObject))
            {
                yield return nested;
            }
        }
    }

    private static void EnsureApplication()
    {
        if (Application.Current is not null)
        {
            return;
        }

        try
        {
            _ = new WinTab.App.App();
        }
        catch (InvalidOperationException)
        {
            // Another test may have already created the single Application instance.
        }
    }

    private static T RunInSta<T>(Func<T> func)
    {
        Exception? exception = null;
        T? result = default;

        var thread = new Thread(() =>
        {
            try
            {
                result = func();
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (exception is not null)
        {
            throw exception;
        }

        return result!;
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
