using System.IO;
using System.Windows.Forms;
using WinTab.UI.Localization;

namespace WinTab.App.Services;

public sealed class TrayIconController : IDisposable
{
    private readonly NotifyIcon _notifyIcon;
    private readonly Action _showSettings;
    private readonly Action _exitApp;
    private bool _disposed;

    public TrayIconController(Action showSettings, Action exitApp)
    {
        _showSettings = showSettings ?? throw new ArgumentNullException(nameof(showSettings));
        _exitApp = exitApp ?? throw new ArgumentNullException(nameof(exitApp));

        var menu = new ContextMenuStrip();

        var openItem = new ToolStripMenuItem(LocalizationManager.GetString("Tray_Open"));
        openItem.Font = new System.Drawing.Font(openItem.Font, System.Drawing.FontStyle.Bold);
        openItem.Click += (_, _) => _showSettings();
        menu.Items.Add(openItem);

        menu.Items.Add(new ToolStripSeparator());

        var exitItem = new ToolStripMenuItem(LocalizationManager.GetString("Tray_Exit"));
        exitItem.Click += (_, _) => _exitApp();
        menu.Items.Add(exitItem);

        _notifyIcon = new NotifyIcon
        {
            Text = LocalizationManager.GetString("Tray_Tooltip"),
            Icon = System.Drawing.SystemIcons.Application,
            Visible = false,
            ContextMenuStrip = menu
        };

        // Try to use the application's own icon
        try
        {
            string? exePath = Environment.ProcessPath;
            if (exePath is not null && File.Exists(exePath))
            {
                var icon = System.Drawing.Icon.ExtractAssociatedIcon(exePath);
                if (icon is not null)
                    _notifyIcon.Icon = icon;
            }
        }
        catch
        {
            // Fall back to system default
        }

        _notifyIcon.DoubleClick += (_, _) => _showSettings();
    }

    public void SetVisible(bool visible)
    {
        _notifyIcon.Visible = visible;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
    }
}
