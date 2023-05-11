using Newtonsoft.Json;

namespace AutoAppenWinform.Models.HideMyAcc
{
    public class HideMyAccBaseRes<T>
    {
        [JsonProperty("code")]
        public string Code { get; set; }

        [JsonProperty("data")]
        public T Data { get; set; }
    }
}