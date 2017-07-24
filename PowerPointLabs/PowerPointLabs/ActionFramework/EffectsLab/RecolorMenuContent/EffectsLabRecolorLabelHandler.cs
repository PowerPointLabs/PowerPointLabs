using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportLabelRibbonId(EffectsLabText.RecolorTag)]
    class EffectsLabRecolorLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            string label = "";
            if (ribbonId.Contains(EffectsLabText.GrayScaleTag))
            {
                label = EffectsLabText.GrayScaleButtonLabel;
            }
            else if (ribbonId.Contains(EffectsLabText.BlackWhiteTag))
            {
                label = EffectsLabText.BlackWhiteButtonLabel;
            }
            else if (ribbonId.Contains(EffectsLabText.GothamTag))
            {
                label = EffectsLabText.GothamButtonLabel;
            }
            else if (ribbonId.Contains(EffectsLabText.SepiaTag))
            {
                label = EffectsLabText.SepiaButtonLabel;
            }
            return label;
        }
    }
}
