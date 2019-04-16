using System;
using System.Collections.Generic;
using System.Speech.Synthesis;
using System.Text;
using System.Text.RegularExpressions;

using PowerPointLabs.TagMatchers;
using PowerPointLabs.Tags;
using PowerPointLabs.Utils;

namespace PowerPointLabs.Models
{
    public class TaggedText
    {
        private const String ClickTagRegex = @"\[AfterClick\]";

        private readonly String _contents;

        public TaggedText(String contents)
        {
            _contents = contents;
        }

        public List<String> SplitByClicks()
        {
            List<String> splitStrings = new List<string>();

            if (String.IsNullOrWhiteSpace(_contents))
            {
                return splitStrings;
            }

            Regex clickRegex = new Regex(ClickTagRegex, RegexOptions.IgnoreCase);
            MatchCollection matches = clickRegex.Matches(_contents);

            int startIndex = 0;
            foreach (Match match in matches)
            {
                String textBefore = _contents.Substring(startIndex, match.Index - startIndex).Trim();
                splitStrings.Add(textBefore);
                startIndex = match.Index + match.Length;
            }

            String remaining = _contents.Substring(startIndex).Trim();
            splitStrings.Add(remaining);

            return splitStrings;
        }

        public PromptBuilder ToPromptBuilder(String defaultVoice)
        {
            PromptBuilder builder = new PromptBuilder(CultureUtil.GetOriginalCulture());
            builder.StartVoice(defaultVoice);

            IEnumerable<ITag> tags = GetTagsInText();

            int startIndex = 0;
            foreach (ITag tag in tags)
            {
                String textBeforeCurrentTag = _contents.Substring(startIndex, tag.Start - startIndex);
                builder.AppendText(textBeforeCurrentTag);

                bool isCommandSuccessful = tag.Apply(builder);
                if (isCommandSuccessful)
                {
                    startIndex = tag.End + 1;
                }
            }

            String remaining = _contents.Substring(startIndex).Trim();
            builder.AppendText(remaining);
            builder.EndVoice();
            return builder;
        }

        public String ToPrettyString()
        {
            StringBuilder builder = new StringBuilder();
            IEnumerable<ITag> tags = GetTagsInText();

            int startIndex = 0;
            foreach (ITag tag in tags)
            {
                String textBeforeCurrentTag = _contents.Substring(startIndex, tag.Start - startIndex);
                builder.Append(textBeforeCurrentTag);
                builder.Append(tag.PrettyPrint());
                startIndex = tag.End + 1;
            }

            String remaining = _contents.Substring(startIndex).Trim();
            builder.Append(remaining);

            return builder.ToString();
        }

        public override string ToString()
        {
            return _contents;
        }

        private IEnumerable<ITag> GetTagsInText()
        {
            List<ITag> tags = new List<ITag>();
            foreach (ITagMatcher tagMatcher in Matchers.All)
            {
                List<ITag> matches = tagMatcher.Matches(_contents);
                tags.AddRange(matches);
            }

            tags.Sort((first, second) => first.Start - second.Start);
            return tags;
        }
    }
}
