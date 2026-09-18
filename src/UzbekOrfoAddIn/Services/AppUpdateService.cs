using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>Explicit metadata check only. Never downloads or executes an installer.</summary>
    public sealed class AppUpdateService
    {
        public const string ReleasesUrl = "https://github.com/nightindex/uzbekorfo/releases";
        public const string LatestApi = "https://api.github.com/repos/nightindex/uzbekorfo/releases/latest";

        public async Task<ReleaseInfo> CheckAsync(CancellationToken cancellation)
        {
            using (var client = new HttpClient(new HttpClientHandler {
                AllowAutoRedirect = false,
                SslProtocols = System.Security.Authentication.SslProtocols.Tls12
            }))
            {
                client.Timeout = TimeSpan.FromSeconds(15);
                client.MaxResponseContentBufferSize = 1024 * 1024;
                client.DefaultRequestHeaders.UserAgent.ParseAdd("UzbekOrfo-UpdateCheck/1.0");
                client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
                client.DefaultRequestHeaders.Add("X-GitHub-Api-Version", "2022-11-28");
                using (var response = await client.GetAsync(LatestApi, cancellation).ConfigureAwait(false))
                {
                    if (response.StatusCode == HttpStatusCode.NotFound) return null;
                    response.EnsureSuccessStatusCode();
                    return Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
                }
            }
        }

        public static ReleaseInfo Parse(string json)
        {
            var data = JObject.Parse(json);
            if (data.Value<bool?>("draft") != false || data.Value<bool?>("prerelease") != false)
                throw new InvalidDataException("Not a stable published release.");
            string tag = data.Value<string>("tag_name") ?? "";
            Version parsed;
            if (!Regex.IsMatch(tag, @"\Av?\d+\.\d+\.\d+(\.\d+)?\z") ||
                !Version.TryParse(tag.TrimStart('v'), out parsed))
                throw new InvalidDataException("Unsupported release version.");
            return new ReleaseInfo {
                Version = Normalize(parsed),
                // Construct from our fixed repository, never trust a response-supplied URL.
                PageUrl = ReleasesUrl + "/tag/" + Uri.EscapeDataString(tag),
                Notes = (data.Value<string>("body") ?? "").Substring(0, Math.Min(12000, (data.Value<string>("body") ?? "").Length))
            };
        }

        public static Version Normalize(Version version) => new Version(version.Major, version.Minor,
            Math.Max(0, version.Build), Math.Max(0, version.Revision));
    }

    public sealed class ReleaseInfo
    {
        public Version Version { get; set; }
        public string PageUrl { get; set; }
        public string Notes { get; set; }
    }
}
