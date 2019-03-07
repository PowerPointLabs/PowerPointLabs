using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        }

        public override string ToString()
        {
            return Voice.ToString();
        }
    }
}
