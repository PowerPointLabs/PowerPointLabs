using System.Speech.Synthesis;

namespace PowerPointLabs.Tags
{
    public class NoEffectTag : Tag
    {
        public NoEffectTag(int start, int end, string contents)
        {
            Start = start;
            End = end;
            Contents = contents;
        }
        public override bool Apply(PromptBuilder builder)
        {
            return true;
        }

        public override string PrettyPrint()
        {
            return "";
        }
    }
}
