using System.Speech.Synthesis;

using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Service;

namespace PowerPointLabs.Tags
{
    public class EndVoiceTag : Tag
    {
        public EndVoiceTag(int start, int end, string contents)
        {
            Start = start;
            End = end;
            Contents = contents;
        }

        public override bool Apply(PromptBuilder builder)
        {
            builder.EndVoice();
            builder.StartVoice(AudioSettingService.selectedVoice.ToString());
            return true;
        }

        public override string PrettyPrint()
        {
            return "";
        }
    }
}
