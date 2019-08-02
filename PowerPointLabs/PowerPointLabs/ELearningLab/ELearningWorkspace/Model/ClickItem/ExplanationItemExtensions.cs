using TestInterface;

namespace PowerPointLabs.ELearningLab.ELearningWorkspace.Model
{
    public static class ExplanationItemExtensions
    {
        public static void CopyFormat(this IExplanationItem self, IExplanationItem template)
        {
            if (self.HasSameFormat(template))
            {
                return;
            }
            self.IsCallout = template.IsCallout;
            self.IsCaption = template.IsCaption;
            self.IsVoice = template.IsVoice;
            self.VoiceLabel = template.VoiceLabel;
            if (template.IsShortVersionIndicated)
            {
                self.CaptionText = template.CaptionText;
                self.IsShortVersionIndicated = true;
            }
            else
            {
                self.IsShortVersionIndicated = false;
            }
        }

        public static bool HasSameFormat(this IExplanationItem self, IExplanationItem other)
        {
            return other == null ||
                (self.IsCallout == other.IsCallout &&
                self.IsCaption == other.IsCaption &&
                self.IsVoice == other.IsVoice &&
                self.VoiceLabel == other.VoiceLabel &&
                self.IsShortVersionIndicated == other.IsShortVersionIndicated);
        }
    }
}
