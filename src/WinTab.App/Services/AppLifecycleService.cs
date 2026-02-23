using System.Runtime.InteropServices;
using WinTab.Core.Enums;
using WinTab.Core.Interfaces;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.Platform.Win32;

namespace WinTab.App.Services;

/// <summary>
/// Orchestrates the startup and shutdown of background services.
/// </summary>
public sealed class AppLifecycleService
{
    private readonly Logger _logger;
    private readonly AppSettings _settings;
    private readonly SettingsStore _settingsStore;
    private readonly SessionStore _sessionStore;
    private readonly IWindowEventSource _windowEventSource;
    private readonly IWindowManager _windowManager;
    private readonly IGroupManager? _groupManager;

    private bool _started;

    public AppLifecycleService(
        Logger logger,
        AppSettings settings,
        SettingsStore settingsStore,
        SessionStore sessionStore,
        IWindowEventSource windowEventSource,
        IWindowManager windowManager,
        IGroupManager? groupManager = null)
    {
        _logger = logger;
        _settings = settings;
        _settingsStore = settingsStore;
        _sessionStore = sessionStore;
        _windowEventSource = windowEventSource;
        _windowManager = windowManager;
        _groupManager = groupManager;
    }

    /// <summary>
    /// Starts all background services based on current settings.
    /// </summary>
    public void Start()
    {
        if (_started) return;
        _started = true;

        _logger.Info("AppLifecycleService starting...");

        _logger.Info("AppLifecycleService started.");
    }

    /// <summary>
    /// Stops all background services and saves session state.
    /// </summary>
    public void Stop()
    {
        if (!_started) return;

        _logger.Info("AppLifecycleService stopping...");

        // Save current session
        SaveSession();

        // Dispose event source
        _windowEventSource.Dispose();

        // Save settings
        try
        {
            _settingsStore.Save(_settings);
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to save settings during shutdown.", ex);
        }

        _started = false;
        _logger.Info("AppLifecycleService stopped.");
    }

    private void SaveSession()
    {
        if (_groupManager is null)
        {
            _logger.Info("No group manager available; skipping session save.");
            return;
        }

        try
        {
            var groups = _groupManager.GetAllGroups();
            var states = groups
                .Select(BuildGroupState)
                .Where(state => state.Tabs.Count >= 2)
                .ToList();

            _sessionStore.SaveSession(states);
            _logger.Info($"Session saved with {states.Count} group(s).");
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to save session.", ex);
        }
    }

    private GroupWindowState BuildGroupState(TabGroup group)
    {
        int activeTabIndex = group.Tabs.FindIndex(t => t.Handle == group.ActiveHandle);

        return new GroupWindowState
        {
            GroupName = group.Name,
            Left = group.Left,
            Top = group.Top,
            Width = group.Width,
            Height = group.Height,
            ActiveTabIndex = activeTabIndex < 0 ? 0 : activeTabIndex,
            State = GetWindowStateMode(group.ActiveHandle),
            Tabs = group.Tabs
                .Select((tab, index) => BuildTabState(tab, index))
                .ToList()
        };
    }

    private GroupWindowTabState BuildTabState(TabItem tab, int order)
    {
        WindowInfo? liveInfo = _windowManager.GetWindowInfo(tab.Handle);

        return new GroupWindowTabState
        {
            Order = order,
            ProcessName = liveInfo?.ProcessName ?? tab.ProcessName,
            WindowTitle = liveInfo?.Title ?? tab.Title,
            ClassName = liveInfo?.ClassName ?? string.Empty,
            ProcessPath = liveInfo?.ProcessPath
        };
    }

    private static GroupWindowStateMode GetWindowStateMode(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
            return GroupWindowStateMode.Normal;

        var placement = new NativeStructs.WINDOWPLACEMENT
        {
            length = (uint)Marshal.SizeOf<NativeStructs.WINDOWPLACEMENT>()
        };

        if (!NativeMethods.GetWindowPlacement(hwnd, ref placement))
            return GroupWindowStateMode.Normal;

        return placement.showCmd switch
        {
            (uint)NativeConstants.SW_SHOWMINIMIZED => GroupWindowStateMode.Minimized,
            (uint)NativeConstants.SW_SHOWMAXIMIZED => GroupWindowStateMode.Maximized,
            _ => GroupWindowStateMode.Normal
        };
    }
}
