using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public static class AzureLocaleToStringConverter
    {
        public static Dictionary<Locale, string> localeMapping = new Dictionary<Locale, string>()
        {
            {Locale.enUS, "en-US"}
        };
    }
}
