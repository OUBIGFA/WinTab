using System.Collections.Generic;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using FluentAssertions;
using WinTab.App.Views.Pages;
using WinTab.Core.Enums;
using WinTab.UI.Localization;
using Xunit;

namespace WinTab.Tests.App;

public class SettingsPagesRuntimeTextRenderingTests
{
    [Fact]
    public void GeneralPage_ShouldRenderMultipleNonEmptyTextBlocks()
    {
        int nonEmptyTextCount = RunInSta(() =>
        {
            EnsureApplication();
            LocalizationManager.ApplyLanguage(Language.Chinese);
            Application.Current!.Resources["General_Startup"].Should().NotBeNull();
            Application.Current!.Resources["General_RunAtStartup"].Should().NotBeNull();

            var page = new GeneralPage(null!);
            return RenderAndCountNonEmptyTextBlocks(page);
        });

        nonEmptyTextCount.Should().BeGreaterThan(6, "General page should render visible text labels and descriptions");
    }

    [Fact]
    public void BehaviorPage_ShouldRenderMultipleNonEmptyTextBlocks()
    {
        int nonEmptyTextCount = RunInSta(() =>
        {
            EnsureApplication();
            LocalizationManager.ApplyLanguage(Language.Chinese);
            Application.Current!.Resources["Behavior_Group_Interception"].Should().NotBeNull();
            Application.Current!.Resources["Behavior_EnableExplorerOpenVerbInterception"].Should().NotBeNull();

            var page = new BehaviorPage(null!);
            return RenderAndCountNonEmptyTextBlocks(page);
        });

        nonEmptyTextCount.Should().BeGreaterThan(6, "Behavior page should render visible text labels and descriptions");
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

    private static int RenderAndCountNonEmptyTextBlocks(Page page)
    {
        var host = new Window
        {
            Width = 1000,
            Height = 700,
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
}
