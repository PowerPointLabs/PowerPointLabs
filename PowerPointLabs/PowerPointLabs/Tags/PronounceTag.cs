using System;
using System.Speech.Synthesis;

namespace PowerPointLabs.Tags
{
    public class PronounceTag : Tag
    {
        public PronounceTag(int start, int end, string contents)
        {
            Start = start;
            End = end;
            Contents = contents;
        }

        public override bool Apply(PromptBuilder builder)
        {
            String pronounciation = ParseTagArgument();

            string wordToPronounce = ParseWordToPronounce();
            try
            {
                builder.AppendTextWithPronunciation(wordToPronounce, pronounciation);
            }
            catch (FormatException)
            {
                return false;
            }
            return true;
        }

        public override string PrettyPrint()
        {
            String word = ParseWordToPronounce();
            return word;
        }

        private string ParseWordToPronounce()
        {
            int firstTagClosingIndex = Contents.IndexOf(']');
            int lastTagStartIndex = Contents.LastIndexOf('[');
            if (lastTagStartIndex - firstTagClosingIndex <= 0)
            {
                return "";
            }
            String wordToPronounce = Contents.Substring(firstTagClosingIndex + 1, lastTagStartIndex - firstTagClosingIndex - 1);
            return wordToPronounce;
        }
    }
}