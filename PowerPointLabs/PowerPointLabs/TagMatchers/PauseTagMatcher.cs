using System.Collections.Generic;
using System.Text.RegularExpressions;

using PowerPointLabs.Tags;

namespace PowerPointLabs.TagMatchers
{
    public class PauseTagMatcher : ITagMatcher
    {
        public Regex Regex { get { return new Regex(@"\[Pause: \d+?(\.?\d+?)?\]", RegexOptions.IgnoreCase); } }
        public List<ITag> Matches(string text)
        {
            List<ITag> foundMatches = new List<ITag>();

            MatchCollection regexMatches = Regex.Matches(text);
            foreach (Match match in regexMatches)
            {
                int matchStart = match.Index;
                int matchEnd = match.Index + match.Length - 1; // 0-based indices.
                PauseTag tag = new PauseTag(matchStart, matchEnd, match.Value);
                foundMatches.Add(tag);
            }

            return foundMatches;
        }
    }
}
