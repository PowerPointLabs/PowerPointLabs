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
                return Voice.ToString();
            }
        }
        public WatsonVoiceType Voice { get; private set; }

        public WatsonVoice(WatsonVoiceType voice)
        {
            Voice = voice;
        }

        public override object Clone()
        {
            WatsonVoice voice = new WatsonVoice(Voice);
            voice.Rank = Rank;
            return voice;
        }

        public override string ToString()
        {
            return Voice.ToString();
        }
    }
}
