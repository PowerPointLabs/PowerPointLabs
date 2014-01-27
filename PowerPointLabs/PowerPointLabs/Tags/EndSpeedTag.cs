using System;
using System.Speech.Synthesis;

namespace AudioGen.Tags
{
    class EndSpeedTag : Tag
    {
        public EndSpeedTag(int start, int end, String contents)
        {
            Start = start;
            End = end;
            Contents = contents;
        }

        public override bool Apply(PromptBuilder builder)
        {
            builder.EndStyle();
            return true;
        }

        public override string PrettyPrint()
        {
            return "";
        }
    }
}
