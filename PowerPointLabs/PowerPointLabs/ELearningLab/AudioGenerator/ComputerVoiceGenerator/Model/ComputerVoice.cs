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
        public ComputerVoice(string voiceName)
        {
            Voice = voiceName;
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
