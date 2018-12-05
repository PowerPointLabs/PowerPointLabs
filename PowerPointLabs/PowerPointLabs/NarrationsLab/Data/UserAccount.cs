using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.NarrationsLab.Data
{
    public class UserAccount
    {
        private static UserAccount instance;

        private string key;
        private string endpoint;

        public static UserAccount GetInstance()
        {
            if (instance == null)
            {
                instance = new UserAccount();
            }
            return instance;
        }

        public void SetUserKeyAndEndpoint(string key, string endpoint)
        {
            this.key = key;
            this.endpoint = endpoint;
        }

        public string GetKey()
        {
            return key;
        }
        public string GetEndpoint()
        {
            return endpoint;
        }

        public string GetUri()
        {
            return EndpointToUriMapping.endpointToUriMapping[endpoint];
        }

        public bool IsEmpty()
        {
            return key == null || endpoint == null;
        }
    }
}
