using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.NarrationsLab.Data
{
    public static class EndpointToUriMapping
    {
        public static Dictionary<string, string> endpointToUriMapping = new Dictionary<string, string>()
        {
            {"https://westus.api.cognitive.microsoft.com/sts/v1.0/issueToken", "https://westus.tts.speech.microsoft.com/cognitiveservices/v1"}
        };
    }
}
