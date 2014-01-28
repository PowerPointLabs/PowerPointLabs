using System;
using System.Speech.Synthesis;

namespace PowerPointLabs.Tags
{
    public interface ITag
    {
        int Start { get; }
        int End { get; }
        String Contents { get; }

        bool Apply(PromptBuilder builder);

        // Returns what this tag should appear like in captions,
        // if anything.
        String PrettyPrint();
    }
}
