using System;
using System.IO;

using PowerPointLabs.ELearningLab.AudioGenerator.WatsonVoiceGenerator.SpeechEngine;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    class SynthesizeWatsonVoice
    {
        private TokenManager tokenManager;
        private WatsonHttpClient client;
        private string endpoint;

        public SynthesizeWatsonVoice(TokenOptions options, string endpoint)
        {
            if (string.IsNullOrEmpty(options.IamApiKey) && string.IsNullOrEmpty(options.IamAccessToken))
            {
                throw new ArgumentNullException(nameof(options.IamAccessToken) + ", "
                    + nameof(options.IamApiKey));
            }

            this.endpoint = endpoint;
            client = new WatsonHttpClient(endpoint);
            tokenManager = new TokenManager(options);
        }
        public MemoryStream Synthesize(Text text, string accept = null, string voice = null)
        {
            if (text == null)
            {
                throw new ArgumentNullException(nameof(text));
            }
            MemoryStream result = null;

            try
            {
                client = client.WithAuthentication(tokenManager.GetToken());

                var restRequest = client.PostAsync($"{endpoint}/v1/synthesize");

                if (!string.IsNullOrEmpty(accept))
                {
                    restRequest.WithHeader("Accept", accept);
                }
                if (!string.IsNullOrEmpty(voice))
                {
                    restRequest.WithArgument("voice", voice);
                }
                restRequest.WithBody<Text>(text);
                restRequest.WithHeader("X-IBMCloud-SDK-Analytics", "service_name=text_to_speech;service_version=v1;operation_id=Synthesize");
                result = new MemoryStream(restRequest.AsByteArray().Result);
            }
            catch (AggregateException ae)
            {
                throw ae.Flatten();
            }

            return result;
        }
    }
}
