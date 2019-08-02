using System.Collections.Generic;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public static class EndpointToUriConverter
    {
        public static Dictionary<string, string> azureRegionToEndpointMapping = new Dictionary<string, string>()
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
        public static Dictionary<string, string> azureEndpointToUriMapping = new Dictionary<string, string>()
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
        public static Dictionary<string, string> watsonRegionToEndpointMapping = new Dictionary<string, string>()
        {
            {"Dallas", "https://stream.watsonplatform.net/text-to-speech/api"},
            { "Washington DC", "https://gateway-wdc.watsonplatform.net/text-to-speech/api"},
            { "Frankfurt", "https://stream-fra.watsonplatform.net/text-to-speech/api"},
            { "Sydney", "https://gateway-syd.watsonplatform.net/text-to-speech/api"},
            { "Tokyo", "https://gateway-tok.watsonplatform.net/text-to-speech/api"},
            { "London", "https://gateway-lon.watsonplatform.net/text-to-speech/api"}
        };
    }
}
