using System.IO.Pipes;
using System.Text;

namespace WinTab.App.Services;

public static class ExplorerOpenRequestClient
{
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
