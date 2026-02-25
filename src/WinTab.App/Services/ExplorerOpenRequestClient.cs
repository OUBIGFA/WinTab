using System.IO.Pipes;
using System.Text;

namespace WinTab.App.Services;

public static class ExplorerOpenRequestClient
{
    /// <summary>
    /// Sends an open-folder request including the foreground window at click time.
    /// The receiver uses this HWND to determine whether the open originated from within Explorer.
    /// </summary>
    public static bool TrySendOpenFolderEx(string path, nint foregroundHwnd, int timeoutMs = 250)
    {
        try
        {
            using var client = new NamedPipeClientStream(
                serverName: ".",
                pipeName: ExplorerOpenRequestServer.PipeName,
                direction: PipeDirection.Out,
                options: PipeOptions.Asynchronous);

            client.Connect(timeoutMs);

            // OPEN_EX <foreground_hwnd_decimal> <path>
            byte[] bytes = Encoding.UTF8.GetBytes($"OPEN_EX {foregroundHwnd} {path}\n");
            client.Write(bytes, 0, bytes.Length);
            client.Flush();
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>Legacy fallback — foreground context is unknown.</summary>
    public static bool TrySendOpenFolder(string path, int timeoutMs = 250)
    {
        try
        {
            using var client = new NamedPipeClientStream(
                serverName: ".",
                pipeName: ExplorerOpenRequestServer.PipeName,
                direction: PipeDirection.Out,
                options: PipeOptions.Asynchronous);

            client.Connect(timeoutMs);

            byte[] bytes = Encoding.UTF8.GetBytes($"OPEN {path}\n");
            client.Write(bytes, 0, bytes.Length);
            client.Flush();
            return true;
        }
        catch
        {
            return false;
        }
    }
}
