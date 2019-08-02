using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Threading.Tasks;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.Extensions;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class Request
    {
        public HttpRequestMessage Message { get; }
        public MediaTypeFormatterCollection Formatters { get; }

        private readonly Lazy<Task<HttpResponseMessage>> dispatch;

        public Request(HttpRequestMessage message, MediaTypeFormatterCollection formatters, Func<Request, Task<HttpResponseMessage>> dispatcher)
        {
            Message = message;
            Formatters = formatters;
            dispatch = new Lazy<Task<HttpResponseMessage>>(() => dispatcher(this));
        }
        public Request WithArgument(string key, string value)
        {
            Message.RequestUri = Message.RequestUri.WithArguments(new KeyValuePair<string, object>(key, value));
            return this;
        }

        public Request WithBody<T>(T text)
        {
            string mediaType = null;
            Message.Content = new ObjectContent<T>(text,
               Formatters.FirstOrDefault(), mediaType);
            return this;
        }

        public Request WithBodyContent(HttpContent body)
        {
            Message.Content = body;
            return this;
        }

        public Request WithHeader(string key, string value)
        {
            if (key == "Accept" && value.StartsWith("audio/", StringComparison.OrdinalIgnoreCase))
            {
                Message.Headers.Accept.Clear();
            }
            Message.Headers.TryAddWithoutValidation(key, value);
            return this;
        }
        public async Task<HttpResponseMessage> AsMessage()
        {
            return await this.GetResponse(this.dispatch.Value).ConfigureAwait(false);
        }

        public async Task<byte[]> AsByteArray()
        {
            HttpResponseMessage message = await this.AsMessage().ConfigureAwait(false);
            return await message.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
        }

        public async Task<T> As<T>()
        {
            HttpResponseMessage message = await this.AsMessage().ConfigureAwait(false);
            return await message.Content.ReadAsAsync<T>(this.Formatters).ConfigureAwait(false);
        }
        private async Task<HttpResponseMessage> GetResponse(Task<HttpResponseMessage> request)
        {           
            try
            {
                HttpResponseMessage response = await request.ConfigureAwait(false);
                return response;
            }
            catch (HttpRequestException e)
            {
                Logger.Log("GetResponse: " + e.InnerException.Message);
                return null;
            }
        }

    }
}
