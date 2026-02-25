using System.Text.Json;
using System.Text.Json.Serialization;
using WinTab.Core.Models;
using WinTab.Diagnostics;

namespace WinTab.Persistence;

/// <summary>
/// Handles loading, saving, and migrating <see cref="AppSettings"/> as JSON on disk.
/// All public members are thread-safe.
/// </summary>
public sealed class SettingsStore : IDisposable
{
    private const int CurrentSchemaVersion = 2;
    private static readonly TimeSpan DebounceInterval = TimeSpan.FromMilliseconds(500);

    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = null, // PascalCase
        Converters = { new JsonStringEnumConverter() },
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly string _settingsPath;
    private readonly Logger? _logger;
    private readonly object _lock = new();
    private Timer? _debounceTimer;
    private AppSettings? _pendingSave;
    private bool _disposed;

    /// <summary>
    /// Creates a new <see cref="SettingsStore"/>.
    /// </summary>
    /// <param name="settingsPath">Absolute path to the settings JSON file.</param>
    /// <param name="logger">Optional logger for diagnostics.</param>
    public SettingsStore(string settingsPath, Logger? logger = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(settingsPath);
        _settingsPath = settingsPath;
        _logger = logger;
    }

    /// <summary>
    /// Loads settings from disk. Returns default settings when the file is
    /// missing or contains invalid JSON.
    /// </summary>
    public AppSettings Load()
    {
        lock (_lock)
        {
            try
            {
                if (!File.Exists(_settingsPath))
                {
                    _logger?.Info("Settings file not found; returning defaults.");
                    return new AppSettings();
                }

                string json = File.ReadAllText(_settingsPath);
                AppSettings? settings = JsonSerializer.Deserialize<AppSettings>(json, SerializerOptions);

                if (settings is null)
                {
                    _logger?.Warn("Settings file deserialized to null; returning defaults.");
                    return new AppSettings();
                }

                if (settings.SchemaVersion > CurrentSchemaVersion)
                {
                    _logger?.Warn(
                        $"Settings schema version {settings.SchemaVersion} is newer than " +
                        $"supported version {CurrentSchemaVersion}; returning defaults.");
                    return new AppSettings();
                }

                bool migrated = MigrateIfNeeded(settings);
                if (migrated)
                    Save(settings);

                _logger?.Info($"Settings loaded successfully (schema v{settings.SchemaVersion}).");
                return settings;
            }
            catch (JsonException ex)
            {
                _logger?.Error($"Corrupt settings file: {ex.Message}");
                return new AppSettings();
            }
            catch (IOException ex)
            {
                _logger?.Error($"Failed to read settings file: {ex.Message}");
                return new AppSettings();
            }
        }
    }

    /// <summary>
    /// Persists settings to disk immediately, creating the parent directory if needed.
    /// </summary>
    public void Save(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        lock (_lock)
        {
            try
            {
                string? directory = Path.GetDirectoryName(_settingsPath);
                if (directory is not null && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                    _logger?.Info($"Created settings directory: {directory}");
                }

                settings.SchemaVersion = CurrentSchemaVersion;
                string json = JsonSerializer.Serialize(settings, SerializerOptions);
                File.WriteAllText(_settingsPath, json);

                _logger?.Info("Settings saved successfully.");
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                _logger?.Error($"Failed to save settings: {ex.Message}");
                throw;
            }
        }
    }

    /// <summary>
    /// Queues a save that executes after <see cref="DebounceInterval"/> of inactivity.
    /// Repeated calls within the window reset the timer so only the last
    /// <paramref name="settings"/> snapshot is written.
    /// </summary>
    public void SaveDebounced(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        lock (_lock)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);

            _pendingSave = settings;

            if (_debounceTimer is null)
            {
                _debounceTimer = new Timer(OnDebounceElapsed, null, DebounceInterval, Timeout.InfiniteTimeSpan);
            }
            else
            {
                _debounceTimer.Change(DebounceInterval, Timeout.InfiniteTimeSpan);
            }
        }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        lock (_lock)
        {
            if (_disposed) return;
            _disposed = true;

            // Flush any pending debounced save before disposing.
            if (_pendingSave is not null)
            {
                try
                {
                    Save(_pendingSave);
                }
                catch (Exception ex)
                {
                    _logger?.Error($"Failed to flush pending save during dispose: {ex.Message}");
                }
                finally
                {
                    _pendingSave = null;
                }
            }

            _debounceTimer?.Dispose();
            _debounceTimer = null;
        }
    }

    // ──────────────────────────────────────────────
    //  Private helpers
    // ──────────────────────────────────────────────

    private void OnDebounceElapsed(object? state)
    {
        AppSettings? snapshot;

        lock (_lock)
        {
            snapshot = _pendingSave;
            _pendingSave = null;
        }

        if (snapshot is not null)
        {
            try
            {
                Save(snapshot);
            }
            catch (Exception ex)
            {
                _logger?.Error($"Debounced save failed: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Applies any required migrations to bring <paramref name="settings"/>
    /// up to <see cref="CurrentSchemaVersion"/>.
    /// </summary>
    private bool MigrateIfNeeded(AppSettings settings)
    {
        bool changed = false;

        if (settings.SchemaVersion < 2)
        {
            settings.PersistExplorerOpenVerbInterceptionAcrossExit = false;
            settings.SchemaVersion = 2;
            _logger?.Info("Migrated settings from v1 to v2.");
            changed = true;
        }

        if (settings.SchemaVersion < CurrentSchemaVersion)
        {
            settings.SchemaVersion = CurrentSchemaVersion;
            _logger?.Info($"Settings migrated to schema v{CurrentSchemaVersion}.");
            changed = true;
        }

        return changed;
    }
}
