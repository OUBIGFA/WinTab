using System.Windows;
using System.Threading;
using System.Windows.Controls;
using WinTab.UI.Views;
using WinTab.Helpers;
using WinTab.UI.Localization;

namespace WinTab;

// ReSharper disable once RedundantExtendsListEntry
public partial class App : Application
{
    private Mutex? _mutex;

    protected override void OnStartup(StartupEventArgs e)
    {
        _mutex = new Mutex(true, Constants.MutexId, out var createdNew);

        if (createdNew)
        {
            base.OnStartup(e);
            LocalizationManager.Initialize();
            SetupTooltipBehavior();

            _ = new MainWindow();
            return;
        }

                CustomMessageBox.Show(
                        LocalizationManager.CurrentLanguage == "zh-CN"
                                ? """
                                    已有实例正在运行。
                                    请在系统托盘图标中查看。
                                    """
                                : """
                                    Another instance is already running.
                                    Check in System Tray Icons.
                                    """,
                        Constants.AppName,
                        icon: MessageBoxImage.Information);
        Shutdown();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _mutex?.Dispose();
        base.OnExit(e);
    }

    private static void SetupTooltipBehavior()
    {
        ToolTipService.ShowDurationProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(3500));
        ToolTipService.InitialShowDelayProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(1700));
        ToolTipService.BetweenShowDelayProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(150));
        ToolTipService.ShowsToolTipOnKeyboardFocusProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(false));
    }
}

