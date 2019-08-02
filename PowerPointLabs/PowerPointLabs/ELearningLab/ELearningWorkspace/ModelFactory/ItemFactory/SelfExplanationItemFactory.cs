using System.Collections.Generic;
using System.Linq;

using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ELearningLab.ELearningWorkspace.ModelFactory
{
    public class SelfExplanationItemFactory : AbstractItemFactory
    {
        public SelfExplanationItemFactory(IEnumerable<ELLEffect> effects) : base(effects)
        { }
        protected override ClickItem CreateBlock()
        {
            if (effects.Count() == 0)
            {
                return null;
            }

            ExplanationItem selfExplanation = new ExplanationItem(captionText: string.Empty);
            foreach (ELLEffect effect in effects)
            {
                string shapeName = effect.shapeName;
                string functionMatch = StringUtility.ExtractFunctionFromString(shapeName);
                selfExplanation.tagNo = SelfExplanationTagService.ExtractTagNo(shapeName);
                switch (functionMatch)
                {
                    case ELearningLabText.CaptionIdentifier:
                        selfExplanation.IsCaption = true;
                        break;
                    case ELearningLabText.CalloutIdentifier:
                        selfExplanation.IsCallout = true;
                        break;
                    case ELearningLabText.AudioIdentifier:
                        selfExplanation.IsVoice = true;
                        selfExplanation.VoiceLabel = StringUtility.ExtractVoiceNameFromString(shapeName);
                        if (StringUtility.ExtractDefaultLabelFromVoiceLabel(selfExplanation.VoiceLabel)
                            .Equals(ELearningLabText.DefaultAudioIdentifier))
                        {
                            selfExplanation.VoiceLabel = string.Format(ELearningLabText.AudioDefaultLabelFormat,
                                AudioSettingService.selectedVoice.ToString());
                        }
                        break;
                    default:
                        break;
                }
            }
            return selfExplanation;
        }
    }
}
