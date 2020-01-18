
using Newtonsoft.Json;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class IamTokenData
    {
        [JsonProperty("access_token", NullValueHandling = NullValueHandling.Ignore)]
        public string AccessToken { get; set; }
        [JsonProperty("refresh_token", NullValueHandling = NullValueHandling.Ignore)]
        public string RefreshToken { get; set; }
        [JsonProperty("token_type", NullValueHandling = NullValueHandling.Ignore)]
        public string TokenType { get; set; }
        [JsonProperty("expires_in", NullValueHandling = NullValueHandling.Ignore)]
        public long? ExpiresIn { get; set; }
        [JsonProperty("expiration", NullValueHandling = NullValueHandling.Ignore)]
        public long? Expiration { get; set; }
    }
}
