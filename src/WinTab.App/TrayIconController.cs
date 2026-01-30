using System.Windows.Forms;
using WinTab.App.Localization;

namespace WinTab.App;

public sealed class TrayIconController : IDisposable
{
    private readonly NotifyIcon _notifyIcon;
    private readonly Action _showSettings;
    private readonly Action _exitApp;

    public TrayIconController(Action showSettings, Action exitApp)
    {
        _showSettings = showSettings;
        _exitApp = exitApp;

        var menu = new ContextMenuStrip();
        menu.Items.Add(new ToolStripMenuItem("打开 WinTab", null, (_, __) => _showSettings()));
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(new ToolStripMenuItem("退出", null, (_, __) => _exitApp()));

        _notifyIcon = new NotifyIcon
        {
            Text = "WinTab - 窗口标签管理",
            Icon = System.Drawing.SystemIcons.Application,
            Visible = true,
            ContextMenuStrip = menu
        };

        _notifyIcon.DoubleClick += (_, __) => _showSettings();
    }

    public void SetVisible(bool visible)
    {
        _notifyIcon.Visible = visible;
    }

    public void Dispose()
    {
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
    }
}
