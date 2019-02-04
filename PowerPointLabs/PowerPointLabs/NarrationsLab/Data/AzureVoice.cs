using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.NarrationsLab.Data
{
    public class AzureVoice
    {
        public Gender voiceType;
        public Locale locale;
        public string voiceName;
        public Voice Voice { get; set; }
        public string Locale
        {
            get
            {
                return LocaleMapping.localeMapping[locale];
            }
        }
        private const string DEFAULTNAMESPACE = "Microsoft Server Speech Text to Speech Voice ";
        private const string LEFTBRACKET = "(";
        private const string RIGHTBRACKET = ")";
        private const string COMMA = ",";
        private const string SPACE = " ";

        public AzureVoice(Gender gender, Locale locale, Voice voice)
        {
            voiceType = gender;
            this.locale = locale;
            this.Voice = voice;
            voiceName = DEFAULTNAMESPACE + LEFTBRACKET + 
                LocaleMapping.localeMapping[locale] + COMMA + SPACE + voice + RIGHTBRACKET;
        }
    }
}
