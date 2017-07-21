using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportLabelRibbonId(TextCollection1.EffectsLabRecolorTag)]
    class EffectsLabRecolorLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            string label = "";
            if (ribbonId.Contains(TextCollection1.GrayScaleTag))
            {
                label = EffectsLabText.GrayScaleButtonLabel;
            }
            else if (ribbonId.Contains(TextCollection1.BlackWhiteTag))
            {
                label = EffectsLabText.BlackWhiteButtonLabel;
            }
            else if (ribbonId.Contains(TextCollection1.GothamTag))
            {
                label = EffectsLabText.GothamButtonLabel;
            }
            else if (ribbonId.Contains(TextCollection1.SepiaTag))
            {
                label = EffectsLabText.SepiaButtonLabel;
            }
            return label;
        }
    }
}
