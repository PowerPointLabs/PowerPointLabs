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
                int percentage = 0;
                if (ribbonId.StartsWith(TextCollection.EffectsLabBlurrinessFeatureSelected))
                {
                    percentage = EffectsLabBlurSelected.CustomPercentageSelected;
                }
                else if (ribbonId.StartsWith(TextCollection.EffectsLabBlurrinessFeatureRemainder))
                {
                    percentage = EffectsLabBlurSelected.CustomPercentageRemainder;
                }
                else if (ribbonId.StartsWith(TextCollection.EffectsLabBlurrinessFeatureBackground))
                {
                    percentage = EffectsLabBlurSelected.CustomPercentageBackground;
                }
                return TextCollection.EffectsLabBlurrinessCustomPrefixLabel + " (" + percentage + "%)";
            }

            int startIndex = ribbonId.IndexOf(TextCollection.DynamicMenuOptionId) + TextCollection.DynamicMenuOptionId.Length;
            string percentageText = ribbonId.Substring(startIndex, ribbonId.Length - startIndex);

            return percentageText + "% " + TextCollection.EffectsLabBlurrinessTag;
        }
    }
}
