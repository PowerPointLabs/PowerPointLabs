using TestInterface;

namespace PowerPointLabs.ELearningLab.ELearningWorkspace.Model
{
    public class ExplanationItemTemplate : IExplanationItem
    {
        public bool IsCallout { get; set; }
        public bool IsCaption { get; set; }
        public bool IsVoice { get; set; }
        public string VoiceLabel { get; set; }
        public bool HasShortVersion { get; set; }
        public string CaptionText { get; set; }
    }
}
