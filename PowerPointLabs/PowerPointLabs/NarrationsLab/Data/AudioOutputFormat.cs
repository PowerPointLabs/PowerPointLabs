using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.NarrationsLab.Data
{
    public enum AudioOutputFormat
    {
        Raw8Khz8BitMonoMULaw,
        Raw16Khz16BitMonoPcm,
        Riff8Khz8BitMonoMULaw,
        Riff16Khz16BitMonoPcm,
        Ssml16Khz16BitMonoSilk,
        Raw16Khz16BitMonoTrueSilk,
        Ssml16Khz16BitMonoTts,
        Audio16Khz128KBitRateMonoMp3,
        Audio16Khz64KBitRateMonoMp3,
        Audio16Khz32KBitRateMonoMp3,
        Audio16Khz16KbpsMonoSiren,
        Riff16Khz16KbpsMonoSiren,
        Raw24Khz16BitMonoTrueSilk,
        Raw24Khz16BitMonoPcm,
        Riff24Khz16BitMonoPcm,
        Audio24Khz48KBitRateMonoMp3,
        Audio24Khz96KBitRateMonoMp3,
        Audio24Khz160KBitRateMonoMp3
    }
}
