using System.Collections.Generic;
using System.Text.RegularExpressions;
using AudioGen.Tags;

namespace AudioGen.TagMatchers
{
    public class PauseTagMatcher : ITagMatcher
    {
        public Regex Regex { get { return new Regex(@"\[Pause: \d+?(\.?\d+?)?\]", RegexOptions.IgnoreCase); } }
        public List<ITag> Matches(string text)
        {
            var foundMatches = new List<ITag>();

            MatchCollection regexMatches = Regex.Matches(text);
            foreach (Match match in regexMatches)
            {
                var matchStart = match.Index;
                var matchEnd = match.Index + match.Length - 1; // 0-based indices.
                PauseTag tag = new PauseTag(matchStart, matchEnd, match.Value);
                foundMatches.Add(tag);
            }

            return foundMatches;
        }
    }
}
