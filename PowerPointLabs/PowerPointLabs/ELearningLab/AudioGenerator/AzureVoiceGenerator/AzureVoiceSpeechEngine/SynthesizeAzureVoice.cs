using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Linq;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class SynthesizeAzureVoice
    {

        private HttpClient client;
        private HttpClientHandler handler;

        public SynthesizeAzureVoice()
        {
            var cookieContainer = new CookieContainer();
            handler = new HttpClientHandler() { CookieContainer = new CookieContainer(), UseProxy = false };
            client = new HttpClient(handler);
        }

        public SynthesizeAzureVoice(HttpClient client)
        {
            this.client = client;
        }

        ~SynthesizeAzureVoice()
        {
            client.Dispose();
            handler.Dispose();
        }

        public event EventHandler<GenericEventArgs<Stream>> OnAudioAvailable;

        public event EventHandler<GenericEventArgs<Exception>> OnError;

        public Task Speak(CancellationToken cancellationToken, InputOptions inputOptions, string filepath)
        {
            client.DefaultRequestHeaders.Clear();
            foreach (var header in inputOptions.Headers)
            {
                client.DefaultRequestHeaders.TryAddWithoutValidation(header.Key, header.Value);
            }

            var genderValue = "";
            switch (inputOptions.VoiceType)
            {
                case Gender.Male:
                    genderValue = "Male";
                    break;

                case Gender.Female:
                default:
                    genderValue = "Female";
                    break;
            }

            Tuple<Task<HttpResponseMessage>, HttpRequestMessage> tuple = SendAsyncHttpRequest(inputOptions, genderValue, cancellationToken);
            var httpTask = tuple.Item1;
            var request = tuple.Item2;
            try
            {
                var result = httpTask.Result;
            }
            catch
            {
                return Task.FromResult(false);
            }

            for (int i = 0; i < 3 && !httpTask.Result.IsSuccessStatusCode; i++)
            {
                MessageBox.Show("Too many requests, please wait for 20 seconds.");
                Thread.Sleep(20000);
                tuple = SendAsyncHttpRequest(inputOptions, genderValue, cancellationToken);
                httpTask = tuple.Item1;
                request = tuple.Item2;
                try
                {
                    var result = httpTask.Result;
                }
                catch 
                {
                    return Task.FromResult(false); 
                }
            }

            var saveTask = httpTask.ContinueWith(
                async (responseMessage, token) =>
                {
                    try
                    {
                        if (responseMessage.IsCompleted && responseMessage.Result != null && responseMessage.Result.IsSuccessStatusCode)
                        {
                            var httpStream = await responseMessage.Result.Content.ReadAsStreamAsync().ConfigureAwait(false);
                            this.AudioAvailable(new GenericEventArgs<Stream>(httpStream, filepath));
                        }
                        else
                        {
                            this.Error(new GenericEventArgs<Exception>(new Exception(String.Format("Service returned {0}", responseMessage.Result.StatusCode))));
                        }
                    }
                    catch (Exception e)
                    {
                        this.Error(new GenericEventArgs<Exception>(e.GetBaseException()));
                    }
                    finally
                    {
                        responseMessage.Dispose();
                        request.Dispose();
                    }
                },
                TaskContinuationOptions.AttachedToParent,
                cancellationToken);

            return saveTask;
        }

        private Tuple<Task<HttpResponseMessage>, HttpRequestMessage> SendAsyncHttpRequest(InputOptions inputOptions, string genderValue, CancellationToken cancellationToken)
        {
            var request = new HttpRequestMessage(HttpMethod.Post, inputOptions.RequestUri)
            {
                Content = new StringContent(GenerateSsml(inputOptions.Locale, genderValue, inputOptions.VoiceName, inputOptions.Text))
            };

            var httpTask = client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);

            return Tuple.Create(httpTask, request);
        }

        private string GenerateSsml(string locale, string gender, string name, string text)
        {
            var ssmlDoc = new XDocument(
                              new XElement("speak",
                                  new XAttribute("version", "1.0"),
                                  new XAttribute(XNamespace.Xml + "lang", "en-US"),
                                  new XElement("voice",
                                      new XAttribute(XNamespace.Xml + "lang", locale),
                                      new XAttribute(XNamespace.Xml + "gender", gender),
                                      new XAttribute("name", name),
                                      text)));
            return ssmlDoc.ToString();
        }

        private void AudioAvailable(GenericEventArgs<Stream> e)
        {
            EventHandler<GenericEventArgs<Stream>> handler = this.OnAudioAvailable;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        private void Error(GenericEventArgs<Exception> e)
        {
            EventHandler<GenericEventArgs<Exception>> handler = this.OnError;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
