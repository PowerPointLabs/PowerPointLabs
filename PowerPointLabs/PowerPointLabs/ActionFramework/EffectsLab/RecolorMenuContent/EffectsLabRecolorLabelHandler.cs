using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.EffectsLab;

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
                label = TextCollection1.EffectsLabGrayScaleButtonLabel;
            }
            else if (ribbonId.Contains(TextCollection1.BlackWhiteTag))
            {
                label = TextCollection1.EffectsLabBlackWhiteButtonLabel;
            }
            else if (ribbonId.Contains(TextCollection1.GothamTag))
            {
                label = TextCollection1.EffectsLabGothamButtonLabel;
            }
            else if (ribbonId.Contains(TextCollection1.SepiaTag))
            {
                label = TextCollection1.EffectsLabSepiaButtonLabel;
            }
            return label;
        }
    }
}
