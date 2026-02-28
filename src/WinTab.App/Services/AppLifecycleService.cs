using System.IO;
using WinTab.Core.Interfaces;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Persistence;

namespace WinTab.App.Services;

/// <summary>
/// Orchestrates the startup and shutdown of background services.
/// </summary>
public sealed class AppLifecycleService
{
    private readonly Logger _logger;
    private readonly AppSettings _settings;
    private readonly SettingsStore _settingsStore;
    private readonly IWindowEventSource _windowEventSource;

    private bool _started;

    public AppLifecycleService(
        Logger logger,
        AppSettings settings,
        SettingsStore settingsStore,
        IWindowEventSource windowEventSource)
    {
        _logger = logger;
        _settings = settings;
        _settingsStore = settingsStore;
        _windowEventSource = windowEventSource;
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
    /// Stops all background services and persists settings.
    /// </summary>
    public void Stop()
    {
        if (!_started) return;

        _logger.Info("AppLifecycleService stopping...");

        // Dispose event source
        _windowEventSource.Dispose();

        // Save settings
        try
        {
            _settingsStore.Save(_settings);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
        {
            _logger.Error("Failed to save settings during shutdown.", ex);
        }

        _started = false;
        _logger.Info("AppLifecycleService stopped.");
    }
}
