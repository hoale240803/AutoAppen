using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System.Text.Json.Serialization;

namespace AutoAppenWinform.Models.HideMyAcc
{
    public class HideMyAccProfile
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonPropertyName("browserType")]
        public string BrowserType { get; set; }

        [JsonProperty("createdAt")]
        public DateTime CreatedAt { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("notes")]
        public string Notes { get; set; }

        [JsonProperty("proxy")]
        public HideMyAccProxy Proxy { get; set; }

        [JsonProperty("os")]
        public string Os { get; set; }

        [JsonProperty("platform")]
        public string Platform { get; set; }
    }

    public class HideMyAccProxy
    {
        [JsonProperty("autoProxyPassword")]
        public string AutoProxyPassword { get; set; }

        [JsonProperty("autoProxyRegion")]
        public string AutoProxyRegion { get; set; }

        [JsonProperty("autoProxyServer")]
        public string AutoProxyServer { get; set; }

        [JsonProperty("autoProxyUsername")]
        public string AutoProxyUsername { get; set; }

        [JsonProperty("host")]
        public string Host { get; set; }

        [JsonProperty("password")]
        public string Password { get; set; }

        [JsonProperty("username")]
        public string Username { get; set; }

        [JsonProperty("mode")]
        public string Mode { get; set; }

        [JsonProperty("port")]
        public string Port { get; set; }

        [JsonProperty("proxyEnabled")]
        public bool ProxyEnabled { get; set; }

        [JsonProperty("torProxyRegion")]
        public string TorProxyRegion { get; set; }
    }

    public class HideMyAccRunProfile
    {
        [JsonProperty("port")]
        public string Port { get; set; }

        [JsonProperty("wsUrl")]
        public string WsUrl { get; set; }
    }
}