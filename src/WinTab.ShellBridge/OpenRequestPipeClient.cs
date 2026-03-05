using System.IO.Pipes;
using System.Text;

namespace WinTab.ShellBridge;

internal static class OpenRequestPipeClient
{
    private const string PipeName = "WinTab_ExplorerOpenRequest";
    private const int DefaultConnectTimeoutMs = 800;
    private const int RetryConnectTimeoutMs = 1200;
    private const int RetryDelayMs = 80;

    public static bool TrySendOpenFolderEx(string path, nint foregroundHwnd)
    {
        if (TrySendOpenFolderExCore(path, foregroundHwnd, DefaultConnectTimeoutMs))
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
