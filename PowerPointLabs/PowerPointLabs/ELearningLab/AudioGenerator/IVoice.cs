using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public abstract class IVoice: ICloneable
    {
        public abstract string VoiceName { get; }
        public int Rank
        {
            get
            {
                return rank;
            }
            set
            {
                rank = value;
            }
        }
        private int rank;
        public abstract object Clone();
    }
}
