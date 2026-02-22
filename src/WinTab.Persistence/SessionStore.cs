using System.Text.Json;
using System.Text.Json.Serialization;
using WinTab.Core.Models;
using WinTab.Diagnostics;

namespace WinTab.Persistence;

/// <summary>
/// Persists and restores a list of <see cref="GroupWindowState"/> objects for
/// crash-recovery and session-restore scenarios.
/// </summary>
public sealed class SessionStore
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = null, // PascalCase
        Converters = { new JsonStringEnumConverter() },
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly string _sessionPath;
    private readonly Logger? _logger;
    private readonly object _lock = new();

    /// <summary>
    /// Creates a new <see cref="SessionStore"/>.
    /// </summary>
    /// <param name="sessionPath">Absolute path to the session backup JSON file.</param>
    /// <param name="logger">Optional logger for diagnostics.</param>
    public SessionStore(string sessionPath, Logger? logger = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sessionPath);
        _sessionPath = sessionPath;
        _logger = logger;
    }

    /// <summary>
    /// Saves the current set of group window states to disk for later recovery.
    /// The parent directory is created automatically if it does not exist.
    /// </summary>
    public void SaveSession(List<GroupWindowState> groups)
    {
        ArgumentNullException.ThrowIfNull(groups);

        lock (_lock)
        {
            try
            {
                string? directory = Path.GetDirectoryName(_sessionPath);
                if (directory is not null && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                    _logger?.Info($"Created session directory: {directory}");
                }

                string json = JsonSerializer.Serialize(groups, SerializerOptions);
                File.WriteAllText(_sessionPath, json);

                _logger?.Info($"Session saved ({groups.Count} group(s)).");
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                _logger?.Error($"Failed to save session: {ex.Message}");
                throw;
            }
        }
    }

    /// <summary>
    /// Loads a previously saved session from disk.
    /// Returns an empty list when the file is missing or contains invalid JSON.
    /// </summary>
    public List<GroupWindowState> LoadSession()
    {
        lock (_lock)
        {
            try
            {
                if (!File.Exists(_sessionPath))
                {
                    _logger?.Info("Session file not found; returning empty list.");
                    return [];
                }

                string json = File.ReadAllText(_sessionPath);
                List<GroupWindowState>? groups = JsonSerializer.Deserialize<List<GroupWindowState>>(json, SerializerOptions);

                if (groups is null)
                {
                    _logger?.Warn("Session file deserialized to null; returning empty list.");
                    return [];
                }

                _logger?.Info($"Session loaded ({groups.Count} group(s)).");
                return groups;
            }
            catch (JsonException ex)
            {
                _logger?.Error($"Corrupt session file: {ex.Message}");
                return [];
            }
            catch (IOException ex)
            {
                _logger?.Error($"Failed to read session file: {ex.Message}");
                return [];
            }
        }
    }

    /// <summary>
    /// Deletes the session backup file if it exists.
    /// </summary>
    public void ClearSession()
    {
        lock (_lock)
        {
            try
            {
                if (File.Exists(_sessionPath))
                {
                    File.Delete(_sessionPath);
                    _logger?.Info("Session file cleared.");
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                _logger?.Error($"Failed to clear session file: {ex.Message}");
                throw;
            }
        }
    }

    /// <summary>
    /// Returns <see langword="true"/> when a session backup file exists on disk.
    /// </summary>
    public bool HasSession()
    {
        lock (_lock)
        {
            return File.Exists(_sessionPath);
        }
    }
}
