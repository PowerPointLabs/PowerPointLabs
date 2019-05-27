using System.Collections.ObjectModel;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public static class AzureVoiceList
    {
        public static ObservableCollection<AzureVoice> voices = new ObservableCollection<AzureVoice>()
        {
            new AzureVoice(Gender.Female, Locale.enUS, AzureVoiceType.ZiraRUS),
            new AzureVoice(Gender.Female, Locale.enUS, AzureVoiceType.JessaRUS),
            new AzureVoice(Gender.Male, Locale.enUS, AzureVoiceType.BenjaminRUS),
            new AzureVoice(Gender.Female, Locale.enUS, AzureVoiceType.Jessa24kRUS),
            new AzureVoice(Gender.Male, Locale.enUS, AzureVoiceType.Guy24kRUS)
        };
    }
}
