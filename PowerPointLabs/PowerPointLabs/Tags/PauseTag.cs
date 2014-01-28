using System;
using System.Speech.Synthesis;

namespace PowerPointLabs.Tags
{
    public class PauseTag : Tag
    {
        private const double TicksPerSecond = 10000000;

        public PauseTag(int start, int end, string contents)
        {
            Start = start;
            End = end;
            Contents = contents;
        }
        public override bool Apply(PromptBuilder builder)
        {
            String argument = ParseTagArgument();

            double duration;
            if (!Double.TryParse(argument, out duration))
            {
                return false;
            }

            int pauseInTicks = (int) Math.Round(duration*TicksPerSecond);

            builder.AppendBreak(new TimeSpan(pauseInTicks));
            return true;
        }

        public override string PrettyPrint()
        {
            return "";
        }
    }
}
