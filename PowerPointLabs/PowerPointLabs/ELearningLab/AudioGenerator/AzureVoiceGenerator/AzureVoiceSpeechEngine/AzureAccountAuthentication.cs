using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class AzureAccountAuthentication
    {
        private string accessUri;
        private string apiKey;
        private string accessToken;
        private Timer accessTokenRenewer;

        //Access token expires every 10 minutes. Renew it every 9 minutes only.
        private const int RefreshTokenDuration = 9;
        private static AzureAccountAuthentication auth = null;

        public static AzureAccountAuthentication GetInstance(string issueTokenUri = null, string apiKey = null)
        {
            if (issueTokenUri != null && apiKey != null)
            {
                auth = new AzureAccountAuthentication(issueTokenUri, apiKey);
            }
            return auth;
        }

        private AzureAccountAuthentication(string issueTokenUri, string apiKey)
        {
            this.accessUri = issueTokenUri;
            this.apiKey = apiKey;
            this.accessToken = HttpPost(issueTokenUri, this.apiKey);

            // renew the token every specfied minutes
            accessTokenRenewer = new Timer(new TimerCallback(OnTokenExpiredCallback),
                                           this,
                                           TimeSpan.FromMinutes(RefreshTokenDuration),
                                           TimeSpan.FromMilliseconds(-1));
        }

        public string GetAccessToken()
        {
            return this.accessToken;
        }

        private void RenewAccessToken()
        {
            string newAccessToken = HttpPost(accessUri, this.apiKey);
            //swap the new token with old one
            //Note: the swap is thread unsafe
            this.accessToken = newAccessToken;
            Console.WriteLine(string.Format("Renewed token for user: {0} is: {1}",
                              this.apiKey,
                              this.accessToken));
        }

        private void OnTokenExpiredCallback(object stateInfo)
        {
            try
            {
                RenewAccessToken();
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Failed renewing access token. Details: {0}", ex.Message));
            }
            finally
            {
                try
                {
                    accessTokenRenewer.Change(TimeSpan.FromMinutes(RefreshTokenDuration), TimeSpan.FromMilliseconds(-1));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(string.Format("Failed to reschedule the timer to renew access token. Details: {0}", ex.Message));
                }
            }
        }

        private string HttpPost(string accessUri, string apiKey)
        {
            // Prepare OAuth request
            HttpWebRequest webRequest = WebRequest.CreateHttp(accessUri);
            webRequest.Method = "POST";
            webRequest.ContentLength = 0;
            webRequest.UserAgent = "PowerPointLabs";
            webRequest.Headers["Ocp-Apim-Subscription-Key"] = apiKey;

            using (WebResponse webResponse = webRequest.GetResponse())
            {
                using (Stream stream = webResponse.GetResponseStream())
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        byte[] waveBytes = null;
                        int count = 0;
                        do
                        {
                            byte[] buf = new byte[1024];
                            count = stream.Read(buf, 0, 1024);
                            ms.Write(buf, 0, count);
                        }
                        while (stream.CanRead && count > 0);

                        waveBytes = ms.ToArray();

                        return Encoding.UTF8.GetString(waveBytes);
                    }
                }
            }
        }
    }
}
