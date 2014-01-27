using System;
using System.Speech.Synthesis;

namespace PowerPointLabs.Tags
{
    class StartSpeedTag : Tag
    {
        public StartSpeedTag(int start, int end, string contents)
        {
            End = end;
            Start = start;
            Contents = contents;
        }

        public override bool Apply(PromptBuilder builder)
        {
            String speed = ParseTagArgument().ToLowerInvariant();
            
            PromptRate rate;
            switch (speed)
            {
                case "fast":
                    rate = PromptRate.Fast;
                    break;
                case "medium":
                    rate = PromptRate.Medium;
                    break;
                case "slow":
                    rate = PromptRate.Slow;
                    break;
                case "extra fast":
                    rate = PromptRate.ExtraFast;
                    break;
                case "extra slow":
                    rate = PromptRate.ExtraSlow;
                    break;
                default:
                    return false;
            }

            PromptStyle style = new PromptStyle(rate);
            builder.StartStyle(style);
            return true;
        }

        public override string PrettyPrint()
        {
            return "";
        }
    }
}
