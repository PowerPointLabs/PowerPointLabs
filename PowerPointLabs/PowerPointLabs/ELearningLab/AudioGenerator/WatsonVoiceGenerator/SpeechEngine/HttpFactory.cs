using System;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Formatting;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    internal static class HttpFactory
    {
        public static HttpRequestMessage GetRequestMessage(HttpMethod method, Uri resource, MediaTypeFormatterCollection formatters)
        {
            HttpRequestMessage request = new HttpRequestMessage(method, resource);

            // add default headers
            request.Headers.Add("accept", formatters.SelectMany(p => p.SupportedMediaTypes).Select(p => p.MediaType));

            return request;
        }
    }
}
