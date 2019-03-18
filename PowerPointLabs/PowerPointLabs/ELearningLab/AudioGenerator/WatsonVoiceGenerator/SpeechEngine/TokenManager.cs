using System;
using System.Collections.Generic;
using System.Net.Http;

using PowerPointLabs.ELearningLab.Extensions;

namespace PowerPointLabs.ELearningLab.AudioGenerator.WatsonVoiceGenerator.SpeechEngine
{
    public class TokenManager
    {
        public WatsonHttpClient Client { get; set; }
        private IamTokenData tokenInfo;
        private string iamUrl;
        private string iamApikey;
        private string userAccessToken;

        public TokenManager(TokenOptions options)
        {
            iamUrl = !string.IsNullOrEmpty(options.IamUrl) ? options.IamUrl
                : "https://iam.bluemix.net/identity/token";
            if (!string.IsNullOrEmpty(options.IamApiKey))
            {
                iamApikey = options.IamApiKey;
            }
            if (!string.IsNullOrEmpty(options.IamAccessToken))
            {
                userAccessToken = options.IamAccessToken;
            }
            tokenInfo = new IamTokenData();
            Client = new WatsonHttpClient(iamUrl);
        }
        public string GetToken()
        {
            if (!string.IsNullOrEmpty(userAccessToken))
            {
                // 1. use user-managed token
                return userAccessToken;
            }
            else if (!string.IsNullOrEmpty(tokenInfo.AccessToken) || IsRefreshTokenExpired())
            {
                // 2. request an initial token
                var tokenInfo = RequestToken();
                SaveTokenInfo(tokenInfo);
                return this.tokenInfo.AccessToken;
            }
            else if (this.IsTokenExpired())
            {
                // 3. refresh a token
                var tokenInfo = RefreshToken();
                SaveTokenInfo(tokenInfo);
                return this.tokenInfo.AccessToken;
            }
            else
            {
                // 4. use valid managed token
                return tokenInfo.AccessToken;
            }
        }
        public void SetAccessToken(string iamAccessToken)
        {
            userAccessToken = iamAccessToken;
        }
        private IamTokenData RequestToken()
        {
            IamTokenData result = null;

            try
            {
                var request = Client.PostAsync(iamUrl);
                request.WithHeader("Content-type", "application/x-www-form-urlencoded");
                request.WithHeader("Authorization", "Basic Yng6Yng=");

                List<KeyValuePair<string, string>> content = new List<KeyValuePair<string, string>>();
                KeyValuePair<string, string> grantType = new KeyValuePair<string, string>("grant_type", "urn:ibm:params:oauth:grant-type:apikey");
                KeyValuePair<string, string> responseType = new KeyValuePair<string, string>("response_type", "cloud_iam");
                KeyValuePair<string, string> apikey = new KeyValuePair<string, string>("apikey", iamApikey);
                content.Add(grantType);
                content.Add(responseType);
                content.Add(apikey);

                var formData = new FormUrlEncodedContent(content);

                request.WithBodyContent(formData);

                result = request.As<IamTokenData>().Result;

                if (result == null)
                {
                    result = new IamTokenData();
                }
            }
            catch (AggregateException ae)
            {
                throw ae.Flatten();
            }

            return result;
        }
        private IamTokenData RefreshToken()
        {
            IamTokenData result = null;

            try
            {
                if (string.IsNullOrEmpty(tokenInfo.RefreshToken))
                {
                    throw new ArgumentNullException(nameof(tokenInfo.RefreshToken));
                }
                var request = Client.PostAsync(iamUrl);
                request.WithHeader("Content-type", "application/x-www-form-urlencoded");
                request.WithHeader("Authorization", "Basic Yng6Yng=");

                List<KeyValuePair<string, string>> content = new List<KeyValuePair<string, string>>();
                KeyValuePair<string, string> grantType = new KeyValuePair<string, string>("grant_type", "refresh_token");
                KeyValuePair<string, string> refreshToken = new KeyValuePair<string, string>("refresh_token", tokenInfo.RefreshToken);
                content.Add(grantType);
                content.Add(refreshToken);

                var formData = new FormUrlEncodedContent(content);

                request.WithBodyContent(formData);

                result = request.As<IamTokenData>().Result;

                if (result == null)
                {
                    result = new IamTokenData();
                }
            }
            catch (AggregateException ae)
            {
                throw ae.Flatten();
            }

            return result;
        }
        private bool IsTokenExpired()
        {
            if (tokenInfo.ExpiresIn == null || tokenInfo.Expiration == null)
            {
                return true;
            }
            float fractionOfTtl = 0.8f;
            long? timeToLive = tokenInfo.ExpiresIn;
            long? expireTime = tokenInfo.Expiration;
            long currentTime = DateTime.Now.ToUnixTimestamp();

            double? refreshTime = expireTime - (timeToLive * (1.0 - fractionOfTtl));
            return refreshTime < currentTime;
        }
        private bool IsRefreshTokenExpired()
        {
            if (tokenInfo.Expiration == null)
            {
                return true;
            };

            long sevenDays = 7 * 24 * 3600;
            long currentTime = DateTime.Now.ToUnixTimestamp();
            long? newTokenTime = tokenInfo.Expiration + sevenDays;
            return newTokenTime < currentTime;
        }
        private void SaveTokenInfo(IamTokenData tokenResponse)
        {
            tokenInfo = tokenResponse;
        }
    }
}
