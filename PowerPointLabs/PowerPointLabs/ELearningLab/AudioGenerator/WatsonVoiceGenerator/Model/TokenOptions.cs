using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class TokenOptions
    {
        private string iamApiKey;
        private string serviceUrl;

        public string IamApiKey
        {
            get { return iamApiKey; }
            set
            {
                iamApiKey = value;
            }
        }

        public string IamAccessToken { get; set; }

        public string ServiceUrl
        {
            get { return serviceUrl; }
            set
            {
                serviceUrl = value;
            }
        }

        public string IamUrl { get; set; }
    }
}
