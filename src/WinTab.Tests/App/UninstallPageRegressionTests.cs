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

public class UninstallPageRegressionTests
{
    [Fact]
    public void UninstallPage_Runtime_ShouldRenderMultipleNonEmptyTextBlocks()
    {
        int nonEmptyTextCount = RunInSta(() =>
        {
            EnsureApplication();
            LocalizationManager.ApplyLanguage(Language.Chinese);

            var page = new UninstallPage(null!);
            return RenderAndCountNonEmptyTextBlocks(page);
        });

        nonEmptyTextCount.Should().BeGreaterThan(5, "Uninstall page should always display rich textual guidance");
    }

    [Fact]
    public void UninstallPage_Design_ShouldAvoidHighSaturationButtonAppearances()
    {
        string pagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", "UninstallPage.xaml");
        var page = System.Xml.Linq.XDocument.Load(pagePath);

        var buttonAppearances = page
            .Descendants()
            .Where(e => e.Name.LocalName == "Button")
            .Select(e => e.Attribute("Appearance")?.Value)
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .ToList();

        buttonAppearances.Should().NotContain("Danger", "uninstall page should keep restrained visual tone");
        buttonAppearances.Should().NotContain("Caution", "uninstall page should keep restrained visual tone");
    }

    [Fact]
    public void UninstallPage_ShouldNotShowDuplicateDeleteUserDataHints()
    {
        string pagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", "UninstallPage.xaml");
        XDocument page = XDocument.Load(pagePath);

        var textKeys = page
            .Descendants()
            .Where(e => e.Name.LocalName == "TextBlock")
            .Select(e => (string?)e.Attribute("Text"))
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .ToList();

        bool hasBottomDefaultHint = textKeys.Contains("{DynamicResource Uninstall_DefaultBehavior}");
        bool hasCheckboxDescHint = textKeys.Contains("{DynamicResource Uninstall_RemoveUserData_Desc}");

        (hasBottomDefaultHint && hasCheckboxDescHint).Should().BeFalse(
            "delete-user-data guidance should only appear once to avoid duplicate copy");
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
