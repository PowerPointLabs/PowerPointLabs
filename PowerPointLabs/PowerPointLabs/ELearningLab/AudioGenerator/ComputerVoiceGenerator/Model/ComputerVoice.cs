using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class ComputerVoice: IVoice
    {
        public string Voice { get; set; }
        public string VoiceName
        {
            get
            {
                return Voice;
            }
        }


        public int Rank
        {
            get
            {
                return rank;
            }
            set
            {
                rank = (int)value;
            }
        }

        private int rank;

        public ComputerVoice(string voiceName)
        {
            Voice = voiceName;
            rank = 0;
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
    }
}
