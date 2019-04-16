using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
