using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Utility;

namespace PowerPointLabs.ELearningLab.Service
{
    public class WatsonRuntimeService
    {
        public static void SaveStringToWaveFile(string text, string filePath, WatsonVoice voice)
        {
            Text synthesizeText = new Text { _Text = text };
            TokenOptions options = new TokenOptions();
            options.IamApiKey = "yZfl6iX33cCGF86qmoyRGXFmerdFPp3kxzePe8OHcvD-";
            options.IamUrl = "https://iam.bluemix.net/identity/token";
            string endpoint = "https://gateway-tok.watsonplatform.net/text-to-speech/api";
            var _service = new SynthesizeWatsonVoice(options, endpoint);
            Logger.Log(voice.VoiceName);

            var synthesizeResult = _service.Synthesize(synthesizeText, "audio/wav", voice: voice.VoiceName);
            StreamUtility.SaveStreamToFile(filePath, synthesizeResult);
        }
    }
}
