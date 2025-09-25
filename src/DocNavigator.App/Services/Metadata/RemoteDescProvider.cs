using System;
using System.Collections.Concurrent;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using DocNavigator.App.Models;

namespace DocNavigator.App.Services.Metadata;

public interface IRemoteDescProvider : IDisposable
{
    Task<string> GetByCodesAsync(DbProfile profile, string serviceCode, string doctypeCode, CancellationToken ct = default);
    void ClearCache();
}

public sealed class RemoteDescProvider : IRemoteDescProvider
{
    private readonly HttpClient _http;
    private readonly ConcurrentDictionary<string, Lazy<Task<string>>> _cache = new();

    public RemoteDescProvider(HttpMessageHandler? handler = null, TimeSpan? timeout = null)
    {
        _http = handler == null ? new HttpClient() : new HttpClient(handler, disposeHandler: false);
        _http.Timeout = timeout ?? TimeSpan.FromSeconds(10);
    }

    public Task<string> GetByCodesAsync(DbProfile profile, string serviceCode, string doctypeCode, CancellationToken ct = default)
    {
        var url = BuildUrl(profile, serviceCode, doctypeCode);
        var lazy = _cache.GetOrAdd(url, u => new Lazy<Task<string>>(() => FetchAsync(u, ct)));
        return lazy.Value;
    }

    private static string BuildUrl(DbProfile profile, string serviceCode, string doctypeCode)
    {
        string Encode(string s) => HttpUtility.UrlEncode(s, Encoding.UTF8);
        var baseUrl = (profile.DescBaseUrl ?? "").TrimEnd('/');
        var template = (profile.DescUrlTemplate ?? "").TrimStart('/');
        var relative = template
            .Replace("{service}", Encode(serviceCode))
            .Replace("{doctype}", Encode(doctypeCode))
            .Replace("{version}", Encode(profile.DescVersion ?? "1.0"));
        return $"{baseUrl}/{relative}";
    }

    private async Task<string> FetchAsync(string url, CancellationToken ct)
    {
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        using var resp = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false);
        if (!resp.IsSuccessStatusCode)
        {
            var body = await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
            throw new HttpRequestException($"Failed to fetch .desc {(int)resp.StatusCode} {resp.ReasonPhrase}\nURL: {url}\n{body}");
        }
        return await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
    }

    public void ClearCache() => _cache.Clear();

    public void Dispose() => _http.Dispose();
}
