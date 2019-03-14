using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public static class AzureEndpointToUriConverter
    {
        public static Dictionary<string, string> regionToEndpointMapping = new Dictionary<string, string>()
        {
            {"West US", "https://westus.api.cognitive.microsoft.com/sts/v1.0/issueToken"},
            { "South East Asia", "https://southeastasia.api.cognitive.microsoft.com/sts/v1.0/issueToken"},
            { "West US2", "https://westus2.api.cognitive.microsoft.com/sts/v1.0/issueToken"},
            { "East US", "https://eastus.api.cognitive.microsoft.com/sts/v1.0/issueToken"},
            { "East US2", "https://eastus2.api.cognitive.microsoft.com/sts/v1.0/issueToken"},
            { "East Asia", "https://eastasia.api.cognitive.microsoft.com/sts/v1.0/issueToken"},
            { "North Europe", "https://northeurope.api.cognitive.microsoft.com/sts/v1.0/issueToken"},
            { "West Europe", "https://westeurope.api.cognitive.microsoft.com/sts/v1.0/issueToken"}
        };
        public static Dictionary<string, string> endpointToUriMapping = new Dictionary<string, string>()
        {
            {"West US", "https://westus.tts.speech.microsoft.com/cognitiveservices/v1"},
            { "South East Asia", "https://southeastasia.tts.speech.microsoft.com/cognitiveservices/v1"},
            { "West US2", "https://westus2.tts.speech.microsoft.com/cognitiveservices/v1"},
            { "East US", "https://eastus.tts.speech.microsoft.com/cognitiveservices/v1"},
            { "East US2", "https://eastus2.tts.speech.microsoft.com/cognitiveservices/v1"},
            { "East Asia", "https://eastasia.tts.speech.microsoft.com/cognitiveservices/v1"},
            { "North Europe", "https://northeurope.tts.speech.microsoft.com/cognitiveservices/v1"},
            { "West Europe", "https://westeurope.tts.speech.microsoft.com/cognitiveservices/v1"}
        };
    }
}
