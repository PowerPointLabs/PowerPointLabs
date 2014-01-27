using System;
using System.Speech.Synthesis;

namespace AudioGen.Tags
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

            var wordToPronounce = ParseWordToPronounce();
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
            String wordToPronounce = Contents.Substring(firstTagClosingIndex, lastTagStartIndex - firstTagClosingIndex);
            return wordToPronounce;
        }
    }
}