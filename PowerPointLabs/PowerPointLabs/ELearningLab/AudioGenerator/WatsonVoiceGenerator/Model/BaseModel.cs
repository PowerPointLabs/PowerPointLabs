using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft.Json;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class BaseModel
    {
        /// <summary>
        /// Custom data object including custom request headers, response headers and response json.
        /// </summary>
        [JsonIgnore]
        public Dictionary<string, object> CustomData
        {
            get
            {
                return _customData == null ? new Dictionary<string, object>() : _customData;
            }
            set
            {
                _customData = value;
            }
        }
        private Dictionary<string, object> _customData;
        /// <summary>
        /// Gets custom request headers.
        /// </summary>
        [JsonIgnore]
        public Dictionary<string, string> CustomRequestHeaders
        {
            get
            {
                return CustomData.ContainsKey(Constants.CUSTOM_REQUEST_HEADERS) ? CustomData[Constants.CUSTOM_REQUEST_HEADERS] as Dictionary<string, string> : null;
            }
        }

        /// <summary>
        /// Gets response headers.
        /// </summary>
        [JsonIgnore]
        public Dictionary<string, string> ResponseHeaders
        {
            get
            {
                return CustomData.ContainsKey(Constants.RESPONSE_HEADERS) ? CustomData[Constants.RESPONSE_HEADERS] as Dictionary<string, string> : null;
            }
        }

        /// <summary>
        /// Gets the response json.
        /// </summary>
        [JsonIgnore]
        public string ResponseJson
        {
            get
            {
                return CustomData.ContainsKey(Constants.JSON) ? CustomData[Constants.JSON] as string : null;
            }
        }
    }
}
