using System.IO.Pipes;
using System.Text;

namespace WinTab.App.Services;

public static class ExplorerOpenRequestClient
{
    private const int DefaultConnectTimeoutMs = 1200;
    private const int RetryConnectTimeoutMs = 1200;
    private const int RetryDelayMs = 80;

    /// <summary>
    /// Sends an open-folder request including the foreground window at click time.
    /// The receiver uses this HWND to determine whether the open originated from within Explorer.
    /// </summary>
    public static bool TrySendOpenFolderEx(string path, nint foregroundHwnd, int timeoutMs = DefaultConnectTimeoutMs)
    {
        if (TrySendOpenFolderExCore(path, foregroundHwnd, timeoutMs))
            return true;

        try
        {
            Thread.Sleep(RetryDelayMs);
        }
        catch
        {
            // ignore
        }

        return TrySendOpenFolderExCore(path, foregroundHwnd, RetryConnectTimeoutMs);
    }

    private static bool TrySendOpenFolderExCore(string path, nint foregroundHwnd, int timeoutMs)
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

}
