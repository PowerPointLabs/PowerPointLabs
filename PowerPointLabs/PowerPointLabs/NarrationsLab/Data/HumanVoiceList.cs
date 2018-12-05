using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.NarrationsLab.Data
{
    public static class HumanVoiceList
    {
        public static ObservableCollection<HumanVoice> voices = new ObservableCollection<HumanVoice>()
        {
            new HumanVoice(Gender.Female, Locale.enUS, Voice.ZiraRUS),
            new HumanVoice(Gender.Female, Locale.enUS, Voice.JessaRUS),
            new HumanVoice(Gender.Male, Locale.enUS, Voice.BenjaminRUS),
            new HumanVoice(Gender.Female, Locale.enUS, Voice.Jessa24kRUS),
            new HumanVoice(Gender.Male, Locale.enUS, Voice.Guy24kRUS)
        };
    }
}
