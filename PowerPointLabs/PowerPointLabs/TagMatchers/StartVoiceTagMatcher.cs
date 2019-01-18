using System.Collections.Generic;
using System.Text.RegularExpressions;

using PowerPointLabs.Tags;

namespace PowerPointLabs.TagMatchers
{
    public class StartVoiceTagMatcher : ITagMatcher
    {
        public Regex Regex { get { return new Regex(@"\[Voice: \w+\]", RegexOptions.IgnoreCase); } }

        public List<ITag> Matches(string text)
        {
            List<ITag> foundMatches = new List<ITag>();

            MatchCollection regexMatches = Regex.Matches(text);
            foreach (Match match in regexMatches)
            {
                int matchStart = match.Index;
                int matchEnd = match.Index + match.Length - 1; // 0-based indices.
                StartVoiceTag tag = new StartVoiceTag(matchStart, matchEnd, match.Value);
                foundMatches.Add(tag);
            }

            return foundMatches;
        }
    }
}