using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportLabelRibbonId(EffectsLabText.BlurrinessTag)]
    class EffectsLabBlurrinessLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            if (ribbonId.Contains(CommonText.DynamicMenuButtonId))
            {
                return EffectsLabText.BlurrinessButtonLabel;
            }

            if (ribbonId.Contains(EffectsLabText.BlurrinessCustom))
            {
                int percentage = 0;
                if (ribbonId.StartsWith(EffectsLabText.BlurrinessFeatureSelected))
                {
                    percentage = EffectsLabSettings.CustomPercentageSelected;
                }
                else if (ribbonId.StartsWith(EffectsLabText.BlurrinessFeatureRemainder))
                {
                    percentage = EffectsLabSettings.CustomPercentageRemainder;
                }
                else if (ribbonId.StartsWith(EffectsLabText.BlurrinessFeatureBackground))
                {
                    percentage = EffectsLabSettings.CustomPercentageBackground;
                }
                return EffectsLabText.BlurrinessCustomPrefixLabel + " (" + percentage + "%)";
            }

            int startIndex = ribbonId.IndexOf(CommonText.DynamicMenuOptionId) + CommonText.DynamicMenuOptionId.Length;
            string percentageText = ribbonId.Substring(startIndex, ribbonId.Length - startIndex);

            return percentageText + "% " + EffectsLabText.BlurrinessTag;
        }
    }
}
