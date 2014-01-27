using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using AudioGen.Tags;

namespace AudioGen.TagMatchers
{
    public interface ITagMatcher
    {
        Regex Regex { get; }

        List<ITag> Matches(String text);
    }
}
