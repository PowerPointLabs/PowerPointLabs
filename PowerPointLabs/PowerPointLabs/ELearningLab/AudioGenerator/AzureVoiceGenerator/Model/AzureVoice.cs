namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class AzureVoice: IVoice
    {
        public Gender voiceType;
        public Locale locale;
        public string voiceName;
        public AzureVoiceType Voice { get; set; }
        public string Locale
        {
            get
            {
                return AzureLocaleToStringConverter.localeMapping[locale];
            }
        }

        public override string VoiceName
        {
            get
            {
                return Voice.ToString();
            }
        }

        private const string DEFAULTNAMESPACE = "Microsoft Server Speech Text to Speech Voice ";
        private const string LEFTBRACKET = "(";
        private const string RIGHTBRACKET = ")";
        private const string COMMA = ",";
        private const string SPACE = " ";

        public AzureVoice(Gender gender, Locale locale, AzureVoiceType voice)
        {
            voiceType = gender;
            this.locale = locale;
            this.Voice = voice;
            voiceName = DEFAULTNAMESPACE + LEFTBRACKET + 
                AzureLocaleToStringConverter.localeMapping[locale] + COMMA + SPACE + voice + RIGHTBRACKET;
            Rank = 0;
        }

        public override string ToString()
        {
            return Voice.ToString();
        }

        public override object Clone()
        {
            AzureVoice voice = new AzureVoice(voiceType, locale, Voice);
            voice.Rank = Rank;
            return voice;
        }
    }
}
