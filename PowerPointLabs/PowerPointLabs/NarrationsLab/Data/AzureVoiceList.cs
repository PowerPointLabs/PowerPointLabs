using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.NarrationsLab.Data
{
    public static class AzureVoiceList
    {
        public static ObservableCollection<AzureVoice> voices = new ObservableCollection<AzureVoice>()
        {
            new AzureVoice(Gender.Female, Locale.enUS, Voice.ZiraRUS),
            new AzureVoice(Gender.Female, Locale.enUS, Voice.JessaRUS),
            new AzureVoice(Gender.Male, Locale.enUS, Voice.BenjaminRUS),
            new AzureVoice(Gender.Female, Locale.enUS, Voice.Jessa24kRUS),
            new AzureVoice(Gender.Male, Locale.enUS, Voice.Guy24kRUS)
        };
    }
}
