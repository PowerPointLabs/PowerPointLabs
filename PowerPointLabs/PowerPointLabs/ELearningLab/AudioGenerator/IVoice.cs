using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public interface IVoice
    {
        string VoiceName { get; }
        int Rank { get; set; }
    }
}
