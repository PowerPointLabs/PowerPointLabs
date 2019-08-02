namespace PowerPointLabs.ELearningLab.AudioGenerator.WatsonVoiceGenerator.Model
{
    public class WatsonAccount
    {
        private static WatsonAccount instance;

        private string key;
        private string region;

        public static WatsonAccount GetInstance()
        {
            if (instance == null)
            {
                instance = new WatsonAccount();
            }
            return instance;
        }

        private WatsonAccount()
        {
            key = null;
            region = null;
        }

        public void SetUserKeyAndRegion(string key, string region)
        {
            this.key = key;
            this.region = region;
        }

        public string GetKey()
        {
            return key;
        }
        public string GetRegion()
        {
            return region;
        }

        public string GetEndpoint()
        {
            return EndpointToUriConverter.watsonRegionToEndpointMapping[region];
        }

        public bool IsEmpty()
        {
            return key == null || region == null;
        }

        public void Clear()
        {
            key = null;
            region = null;
        }
    }
}
