using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class AzureAccount
    {
        private static AzureAccount instance;

        private string key;
        private string endpoint;

        public static AzureAccount GetInstance()
        {
            if (instance == null)
            {
                instance = new AzureAccount();
            }
            return instance;
        }

        private AzureAccount()
        {
            key = null;
            endpoint = null;
        }

        public void SetUserKeyAndRegion(string key, string endpoint)
        {
            this.key = key;
            this.endpoint = endpoint;
        }

        public string GetKey()
        {
            return key;
        }
        public string GetRegion()
        {
            return endpoint;
        }

        public string GetUri()
        {
            if (!string.IsNullOrEmpty(endpoint))
            {
                return EndpointToUriConverter.azureEndpointToUriMapping[endpoint];
            }
            return null;
        }

        public bool IsEmpty()
        {
            return key == null || endpoint == null;
        }

        public void Clear()
        {
            key = null;
            endpoint = null;
        }
    }
}
