using System.IO.Pipes;
using System.Text;

namespace WinTab.ShellBridge;

internal static class OpenRequestPipeClient
{
    private const string PipeName = "WinTab_ExplorerOpenRequest";
    // DelegateExecute runs in explorer.exe. Keep the sync budget tiny
    // to avoid visibly freezing taskbar interactions when the server is unavailable.
    private const int DefaultConnectTimeoutMs = 80;
    private const int RetryConnectTimeoutMs = 120;
    private const int RetryDelayMs = 20;

    public static bool TrySendOpenFolderEx(string path, nint foregroundHwnd, bool allowRetry = true)
    {
        if (TrySendOpenFolderExCore(path, foregroundHwnd, DefaultConnectTimeoutMs))
            return true;

        if (!allowRetry)
            return false;

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
                ".",
                PipeName,
                PipeDirection.Out,
                PipeOptions.Asynchronous);

            client.Connect(timeoutMs);
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
