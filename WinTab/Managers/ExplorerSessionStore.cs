using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Models;

namespace WinTab.Managers;

/// <summary>
/// Stores only the last closed window. Saves are serialized/coalesced off the caller's thread; readers
/// see the latest complete in-memory snapshot immediately, even when the disk is slow or unavailable.
/// </summary>
internal sealed class ExplorerSessionStore : IDisposable
{
    internal const int MaxFileBytes = 4 * 1024 * 1024;
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly object _gate = new();
    private readonly string _path;
    private readonly Action<ExplorerSession> _write;
    private ExplorerSession? _snapshot;
    private Exception? _lastError;
    private TaskCompletionSource<bool>? _completion;
    private long _revision;
    private long _writtenRevision;
    private bool _saving;
    private bool _preserveBackup;
    private bool _saveBlocked;
    private bool _disposed;

    public ExplorerSessionStore(string path, Action<ExplorerSession>? write = null)
    {
        _path = Path.GetFullPath(path);
        _write = write ?? WriteAtomically;
        _snapshot = Load();
    }

    public ExplorerSession? Snapshot
    {
        get { lock (_gate) return _snapshot?.Copy(); }
    }
    public Exception? LastError => Volatile.Read(ref _lastError);
    public event Action? ErrorChanged;

    public Task<bool> SaveAsync(ExplorerSession session)
    {
        var snapshot = session.ValidatedCopy();
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            _snapshot = snapshot;
            _revision++;
            return StartSave();
        }
    }

    public Task<bool> FlushAsync()
    {
        lock (_gate)
            return _disposed ? Task.FromResult(_writtenRevision == _revision && LastError == null) : StartSave();
    }

    private Task<bool> StartSave()
    {
        if (_saveBlocked)
            return Task.FromResult(false);
        if (_saving)
            return _completion!.Task;
        if (_revision == _writtenRevision)
            return Task.FromResult(LastError == null);
        _saving = true;
        _completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        _ = Task.Run(SavePending);
        return _completion.Task;
    }

    private void SavePending()
    {
        while (true)
        {
            ExplorerSession snapshot;
            long revision;
            lock (_gate)
            {
                snapshot = _snapshot!.Copy();
                revision = _revision;
            }
            Exception? error = null;
            try { _write(snapshot); }
            catch (Exception exception) { error = exception; }
            var previousError = Interlocked.Exchange(ref _lastError, error);
            if (error != null || previousError != null)
            {
                try { ErrorChanged?.Invoke(); }
                catch (Exception exception) { Trace.TraceError($"Session storage status notification failed: {exception}"); }
            }
            lock (_gate)
            {
                if (error == null)
                    _writtenRevision = revision;
                if (_revision > revision)
                    continue;
                _saving = false;
                _completion!.TrySetResult(error == null);
                return;
            }
        }
    }

    private ExplorerSession? Load()
    {
        try { return Read(_path); }
        catch (Exception exception) when (IsStorageFailure(exception)) { RecordLoadFailure(exception); }
        try
        {
            var backup = Read(_path + ".bak");
            _preserveBackup = true;
            _lastError ??= new IOException("The session file was missing; the last valid backup was recovered.");
            return backup;
        }
        catch (Exception exception) when (IsStorageFailure(exception))
        {
            RecordLoadFailure(exception);
            return null;
        }
    }

    private void RecordLoadFailure(Exception exception)
    {
        if (exception is FileNotFoundException or DirectoryNotFoundException) return;
        _lastError ??= exception;
        // Corrupt JSON can be replaced, but unknown future formats and unreadable files may contain
        // valid user data. Neither authorizes writing a fallback over that data.
        if (exception is not JsonException || exception is ExplorerSession.UnsupportedVersionException)
            _saveBlocked = true;
    }

    private static bool IsStorageFailure(Exception exception) =>
        exception is IOException or UnauthorizedAccessException or JsonException;

    private static ExplorerSession Read(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (stream.Length > MaxFileBytes)
            throw new IOException("The session file exceeds the supported size; it was not overwritten.");
        using var document = JsonDocument.Parse(stream);
        if (document.RootElement.ValueKind != JsonValueKind.Object)
            throw new JsonException("The session must be an object.");
        if (document.RootElement.TryGetProperty(nameof(ExplorerSession.Version), out var version) &&
            version.ValueKind == JsonValueKind.Number && version.TryGetInt32(out var number) &&
            number != ExplorerSession.CurrentVersion)
            throw new ExplorerSession.UnsupportedVersionException(number);
        return (document.RootElement.Deserialize<ExplorerSession>() ?? throw new JsonException("The session must be an object."))
            .ValidatedCopy();
    }

    private void WriteAtomically(ExplorerSession snapshot)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        var temporary = _path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                JsonSerializer.Serialize(stream, snapshot, JsonOptions);
                if (stream.Length > MaxFileBytes)
                    throw new JsonException("The serialized session exceeds the supported size.");
                stream.Flush(flushToDisk: true);
            }
            if (File.Exists(_path))
            {
                var preserveBackup = _preserveBackup;
                try { _ = Read(_path); }
                catch (ExplorerSession.UnsupportedVersionException) { throw; }
                catch (JsonException) { preserveBackup = true; }
                File.Replace(temporary, _path, preserveBackup ? null : _path + ".bak");
            }
            else
            {
                File.Move(temporary, _path);
            }
            _preserveBackup = false;
        }
        finally
        {
            try { if (File.Exists(temporary)) File.Delete(temporary); }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                Trace.TraceError($"Session temporary file could not be removed: {exception.Message}");
            }
        }
    }

    public void Dispose()
    {
        lock (_gate) _disposed = true;
    }
}
