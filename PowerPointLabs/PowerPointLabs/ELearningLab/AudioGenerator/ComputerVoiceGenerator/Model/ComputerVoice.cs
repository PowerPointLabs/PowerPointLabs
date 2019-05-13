using System.Collections.Generic;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class ComputerVoice: IVoice
    {
        public string Voice { get; set; }
        public override string VoiceName
        {
            get
            {
                return Voice;
            }
        }

        public ComputerVoice(string voiceName)
        {
            Voice = voiceName;
            Rank = 0;
        }
        public override string ToString()
        {
            return Voice;
        }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is ComputerVoice))
            {
                return false;
            }
            return ((ComputerVoice)obj).Voice == Voice;
        }

        public override int GetHashCode()
        {
            return EqualityComparer<string>.Default.GetHashCode(Voice);
        }

        public override object Clone()
        {
            ComputerVoice voice = new ComputerVoice(Voice);
            voice.Rank = Rank;
            return voice;
        }
    }
}
