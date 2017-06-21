using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.EffectsLab;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(TextCollection.EffectsLabBlurrinessTag)]
    class EffectsLabBlurrinessLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            if (ribbonId.Contains(TextCollection.DynamicMenuButtonId))
            {
                return TextCollection.EffectsLabBlurrinessButtonLabel;
            }

            if (ribbonId.Contains(TextCollection.EffectsLabBlurrinessCustom))
            {
                if (ribbonId.StartsWith(TextCollection.EffectsLabBlurrinessFeatureSelected))
                {
                    return EffectsLabBlurSelected.CustomPercentageSelected + "% " + TextCollection.EffectsLabBlurrinessTag;
                }
                else if (ribbonId.StartsWith(TextCollection.EffectsLabBlurrinessFeatureRemainder))
                {
                    return EffectsLabBlurSelected.CustomPercentageRemainder + "% " + TextCollection.EffectsLabBlurrinessTag;
                }
                else if (ribbonId.StartsWith(TextCollection.EffectsLabBlurrinessFeatureBackground))
                {
                    return EffectsLabBlurSelected.CustomPercentageBackground + "% " + TextCollection.EffectsLabBlurrinessTag;
                }
            }

            int startIndex = ribbonId.IndexOf(TextCollection.DynamicMenuOptionId) + TextCollection.DynamicMenuOptionId.Length;
            string percentage = ribbonId.Substring(startIndex, ribbonId.Length - startIndex);

            return percentage + "% " + TextCollection.EffectsLabBlurrinessTag;
        }
    }
}
