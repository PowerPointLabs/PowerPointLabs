using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class WatsonVoice: IVoice
    {
        public override string VoiceName
        {
            get
            {
                return "en-US_" + voice.ToString();
            }
        }
        private WatsonVoiceType voice;

        public WatsonVoice(WatsonVoiceType voice)
        {
            this.voice = voice;
        }

        public override object Clone()
        {
            WatsonVoice voice = new WatsonVoice(this.voice);
            voice.Rank = Rank;
            return voice;
        }
    }
}
