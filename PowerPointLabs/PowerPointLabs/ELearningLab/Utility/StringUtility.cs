using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ELearningLab.Utility
{
    public class StringUtility
    {
        // input str is "PPTL_tag#_function(_voiceName)(_Default)"
        public static string ExtractFunctionFromString(string str)
        {
            Regex regex = new Regex(ELearningLabText.ExtractFunctionRegex, RegexOptions.IgnoreCase);
            Match match = regex.Match(str);
            string value = "";
            if (match.Success)
            {
                value = match.Groups[1].Value.Trim();
            }
            return value;
        }

        // input str is "PPTL_tag#_function(_voiceName)(_Default)"
        public static string ExtractVoiceNameFromString(string str)
        {
            Regex regex = new Regex(ELearningLabText.ExtractVoiceNameRegex, 
                RegexOptions.IgnoreCase);
            Match match = regex.Match(str);
            string value = "";
            if (match.Success)
            {
                value = match.Groups[1].Value.Trim();
            }
            return value;
        }

        // input str is "voiceName(_Default)"
        public static string ExtractVoiceNameFromVoiceLabel(string str)
        {
            return str.Split(ELearningLabText.Underscore.ToCharArray().First())[0];
        }

        public static string ExtractDefaultLabelFromVoiceLabel(string str)
        {
            if (str.Split(ELearningLabText.Underscore.ToCharArray().First()).Count() > 1)
            {
                return str.Split(ELearningLabText.Underscore.ToCharArray().First())[1];
            }
            return string.Empty;
        }

        public static bool IsPPTLShape(string shapeName)
        {
            return shapeName.Contains(ELearningLabText.Identifier);
        }
    }
}
