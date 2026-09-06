using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Managers;

internal static class UpdateManagerTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("an update timeout returns a failure instead of escaping into the UI", ReportsTimeout);
        yield return ("an update deadline also covers a stalled response body", ReportsBodyTimeout);
        yield return ("an explicitly cancelled update check stays cancelled", PreservesCancellation);
        yield return ("malformed update responses report failure", RejectsMalformedResponse);
        yield return ("a valid update response exposes the matching installer", ReadsValidResponse);
    }

    private static async Task ReportsTimeout()
    {
        using var client = new HttpClient(new ResponseHandler(async cancellationToken =>
        {
            await Task.Delay(Timeout.Infinite, cancellationToken);
            return new HttpResponseMessage(HttpStatusCode.OK);
        })) { Timeout = TimeSpan.FromMilliseconds(40) };
        var result = await UpdateManager.CheckForUpdatesWithResultAsync(client).WaitAsync(TimeSpan.FromSeconds(2));
        Check.That(!result.Completed && !string.IsNullOrWhiteSpace(result.ErrorMessage),
            "A network deadline must return an actionable failure without crashing the caller.");
    }

    private static async Task ReportsBodyTimeout()
    {
        using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(1));
        using var client = new HttpClient(new ResponseHandler(token =>
            Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(new PendingReadStream())
            }))) { Timeout = TimeSpan.FromMilliseconds(40) };

        var result = await UpdateManager.CheckForUpdatesWithResultAsync(client, cancellation.Token);
        Check.That(!result.Completed && !string.IsNullOrWhiteSpace(result.ErrorMessage),
            "The request deadline must also stop reading a server response that never completes.");
    }

    private static async Task PreservesCancellation()
    {
        using var cancellation = new CancellationTokenSource();
        using var client = new HttpClient(new ResponseHandler(async token =>
        {
            cancellation.Cancel();
            await Task.Delay(Timeout.Infinite, token);
            return new HttpResponseMessage(HttpStatusCode.OK);
        }));
        try
        {
            await UpdateManager.CheckForUpdatesWithResultAsync(client, cancellation.Token);
        }
        catch (OperationCanceledException) when (cancellation.IsCancellationRequested)
        {
            return;
        }
        throw new InvalidOperationException("Explicit cancellation must not be reported as a network failure.");
    }

    private static async Task RejectsMalformedResponse()
    {
        foreach (var body in new[] { "not json", "{}", "{\"tag_name\":true}", "{\"tag_name\":\"latest\"}" })
        {
            using var client = CreateClient(body);
            Check.That(!(await UpdateManager.CheckForUpdatesWithResultAsync(client)).Completed,
                "Invalid release metadata must not be accepted as an update.");
        }
    }

    private static async Task ReadsValidResponse()
    {
        var architecture = RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant();
        var body = "{\"tag_name\":\"v99.0.0\",\"assets\":[{\"name\":\"WinTab_v99.0.0_" + architecture +
            "_Setup.exe\",\"browser_download_url\":\"https://example.test/setup.exe\"}]}";
        using var client = CreateClient(body);
        var result = await UpdateManager.CheckForUpdatesWithResultAsync(client);
        Check.That(result.Completed && result.UpdateAvailable, "A newer valid release must be detected.");
        Check.Equal("https://example.test/setup.exe", result.DownloadUrl, "Use the compatible installer URL.");
    }

    private static HttpClient CreateClient(string body) => new(new ResponseHandler(token =>
        Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(body) })));

    private sealed class ResponseHandler(Func<CancellationToken, Task<HttpResponseMessage>> response) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
            response(cancellationToken);
    }

    private sealed class PendingReadStream : MemoryStream
    {
        public override bool CanSeek => false;

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            await Task.Delay(Timeout.Infinite, cancellationToken);
            return 0;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
    }
}
