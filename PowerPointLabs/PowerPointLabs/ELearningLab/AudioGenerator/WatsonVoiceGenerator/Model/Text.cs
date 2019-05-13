
using Newtonsoft.Json;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class Text : BaseModel
    {
        /// <summary>
        /// The text to synthesize.
        /// </summary>
        [JsonProperty("text", NullValueHandling = NullValueHandling.Ignore)]
        public string _Text { get; set; }
    }
}
