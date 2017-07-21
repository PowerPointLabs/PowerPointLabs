using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.EffectsLab;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportLabelRibbonId(TextCollection1.EffectsLabBlurrinessTag)]
    class EffectsLabBlurrinessLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            if (ribbonId.Contains(TextCollection1.DynamicMenuButtonId))
            {
                return TextCollection1.EffectsLabBlurrinessButtonLabel;
            }

            if (ribbonId.Contains(TextCollection1.EffectsLabBlurrinessCustom))
            {
                int percentage = 0;
                if (ribbonId.StartsWith(TextCollection1.EffectsLabBlurrinessFeatureSelected))
                {
                    percentage = EffectsLabSettings.CustomPercentageSelected;
                }
                else if (ribbonId.StartsWith(TextCollection1.EffectsLabBlurrinessFeatureRemainder))
                {
                    percentage = EffectsLabSettings.CustomPercentageRemainder;
                }
                else if (ribbonId.StartsWith(TextCollection1.EffectsLabBlurrinessFeatureBackground))
                {
                    percentage = EffectsLabSettings.CustomPercentageBackground;
                }
                return TextCollection1.EffectsLabBlurrinessCustomPrefixLabel + " (" + percentage + "%)";
            }

            int startIndex = ribbonId.IndexOf(TextCollection1.DynamicMenuOptionId) + TextCollection1.DynamicMenuOptionId.Length;
            string percentageText = ribbonId.Substring(startIndex, ribbonId.Length - startIndex);

            return percentageText + "% " + TextCollection1.EffectsLabBlurrinessTag;
        }
    }
}
