using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace WinTab.UI.Desktop;

/// <summary>
/// Reads bounded lines from one pipe, retaining read-ahead for the next frame. Buffering avoids an
/// asynchronous read for every character; checking before appending keeps damaged output bounded.
/// A reader belongs to one stream and must not be used concurrently.
/// </summary>
internal sealed class DesktopLineReader(TextReader reader, bool requireTerminator = true)
{
    private readonly char[] _buffer = new char[4096];
    private int _position;
    private int _count;

    internal async Task<string?> ReadLineAsync()
    {
        var line = new StringBuilder();
        while (true)
        {
            if (_position == _count)
            {
                _count = await reader.ReadAsync(_buffer.AsMemory()).ConfigureAwait(false);
                _position = 0;
                if (_count == 0)
                {
                    if (line.Length == 0) return null;
                    if (requireTerminator) throw new IOException("Incomplete UI message");
                    return line.ToString().TrimEnd('\r');
                }
            }

            var newline = Array.IndexOf(_buffer, '\n', _position, _count - _position);
            var end = newline < 0 ? _count : newline;
            var length = end - _position;
            if (line.Length + length > DesktopCommand.MaximumMessageLength)
                throw new IOException("UI message too large");
            line.Append(_buffer, _position, length);
            _position = newline < 0 ? end : end + 1;
            if (newline >= 0) return line.ToString().TrimEnd('\r');
        }
    }
}
