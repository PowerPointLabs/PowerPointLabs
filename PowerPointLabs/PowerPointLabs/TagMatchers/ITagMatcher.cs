using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using PowerPointLabs.Tags;

namespace PowerPointLabs.TagMatchers
{
    public interface ITagMatcher
    {
        Regex Regex { get; }

        List<ITag> Matches(String text);
    }
}
