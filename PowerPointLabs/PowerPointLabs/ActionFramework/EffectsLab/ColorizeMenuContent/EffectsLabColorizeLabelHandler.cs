using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.EffectsLab;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportLabelRibbonId(TextCollection.EffectsLabColorizeTag)]
    class EffectsLabColorizeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            string label = "";
            if (ribbonId.Contains(TextCollection.GrayScaleTag))
            {
                label = TextCollection.EffectsLabGrayScaleButtonLabel;
            }
            else if (ribbonId.Contains(TextCollection.BlackWhiteTag))
            {
                label = TextCollection.EffectsLabBlackWhiteButtonLabel;
            }
            else if (ribbonId.Contains(TextCollection.GothamTag))
            {
                label = TextCollection.EffectsLabGothamButtonLabel;
            }
            else if (ribbonId.Contains(TextCollection.SepiaTag))
            {
                label = TextCollection.EffectsLabSepiaButtonLabel;
            }
            return label;
        }
    }
}
