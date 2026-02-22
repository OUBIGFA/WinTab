using WinTab.Core.Interfaces;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Platform.Win32;
using WinTab.TabHost;

namespace WinTab.App.Services;

/// <summary>
/// Lazily wires up the heavy tab-host infrastructure (group manager,
/// auto-group engine, drag-to-group handler) only when the user has
/// the related features enabled.
/// </summary>
public sealed class TabHostBootstrapper : IDisposable
{
    private readonly AppSettings _settings;
    private readonly IGroupManager _groupManager;
    private readonly IWindowManager _windowManager;
    private readonly IWindowEventSource _windowEventSource;
    private readonly DragDetector _dragDetector;
    private readonly Logger _logger;

    private readonly object _lock = new();

    private AutoGroupEngine? _autoGroupEngine;
    private DragToGroupHandler? _dragHandler;
    private bool _initialized;

    public TabHostBootstrapper(
        AppSettings settings,
        IGroupManager groupManager,
        IWindowManager windowManager,
        IWindowEventSource windowEventSource,
        DragDetector dragDetector,
        Logger logger)
    {
        _settings = settings;
        _groupManager = groupManager;
        _windowManager = windowManager;
        _windowEventSource = windowEventSource;
        _dragDetector = dragDetector;
        _logger = logger;
    }

    /// <summary>
    /// Returns <c>true</c> when any tab-host feature is enabled in settings.
    /// </summary>
    public bool RequiresInitialization =>
        _settings.EnableDragToGroup || _settings.AutoApplyRules;

    /// <summary>
    /// Initializes the tab-host services if they have not already been started.
    /// Safe to call multiple times.
    /// </summary>
    public void EnsureInitialized()
    {
        if (!RequiresInitialization)
        {
            _logger.Info("Tab host features disabled; skipping bootstrap.");
            return;
        }

        lock (_lock)
        {
            if (_initialized)
                return;

            if (_settings.AutoApplyRules)
            {
                _autoGroupEngine = new AutoGroupEngine(
                    _groupManager,
                    _windowEventSource,
                    _windowManager,
                    () => _settings);
                _autoGroupEngine.Start();
                _logger.Info("AutoGroupEngine started.");
            }

            if (_settings.EnableDragToGroup)
            {
                _dragHandler = new DragToGroupHandler(_groupManager, _windowManager);
                if (!_dragDetector.IsEnabled)
                    _dragDetector.Enable();
                _dragHandler.Enable(_dragDetector);
                _logger.Info("Drag-to-group handler enabled.");
            }

            _initialized = true;
            _logger.Info("Tab host services initialized.");
        }
    }

    public void Shutdown()
    {
        lock (_lock)
        {
            _autoGroupEngine?.Stop();
            _autoGroupEngine?.Dispose();
            _autoGroupEngine = null;

            _dragHandler?.Disable();
            _dragHandler?.Dispose();
            _dragHandler = null;

            _initialized = false;
        }
    }

    public void Dispose()
    {
        Shutdown();
    }
}
