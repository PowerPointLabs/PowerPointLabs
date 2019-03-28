using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;

namespace PowerPointLabs.ELearningLab.AudioGenerator.WatsonVoiceGenerator.SpeechEngine
{
    public class WatsonHttpClient
    {
        public HttpClient BaseClient { get; set; }
        public MediaTypeFormatterCollection Formatters { get; protected set; }
        public WatsonHttpClient(string baseUri)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            BaseClient = new HttpClient();
            if (baseUri != null)
            {
                BaseClient.BaseAddress = new Uri(baseUri);
            }
            Formatters = new MediaTypeFormatterCollection();
        }

        public WatsonHttpClient()
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            BaseClient = new HttpClient();
            Formatters = new MediaTypeFormatterCollection();
        }

        public WatsonHttpClient WithAuthentication(string apiToken)
        {
            if (!string.IsNullOrEmpty(apiToken))
            {
                this.BaseClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiToken);
            }

            return this;
        }

        public Request PostAsync(string resource)
        {
            return SendAsync(HttpMethod.Post, resource);
        }
        public virtual Request SendAsync(HttpMethod method, string resource)
        {
            Uri uri = new Uri(this.BaseClient.BaseAddress, resource);
            HttpRequestMessage message = HttpFactory.GetRequestMessage(method, uri, Formatters);
            return this.SendAsync(message);
        }
        public virtual Request SendAsync(HttpRequestMessage message)
        {
            return new Request(message, Formatters, request => BaseClient.SendAsync(request.Message));
        }
    }
}
