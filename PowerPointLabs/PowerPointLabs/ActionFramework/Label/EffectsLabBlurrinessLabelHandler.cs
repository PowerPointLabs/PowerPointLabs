using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

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

            if (ribbonId.Contains(TextCollection.DynamicMenuCheckBoxId))
            {
                var checkBoxStartIndex = ribbonId.IndexOf("Blur") + 4;
                var length = ribbonId.IndexOf(TextCollection.DynamicMenuCheckBoxId) - checkBoxStartIndex;
                var checkBoxFeatureLabel = ribbonId.Substring(checkBoxStartIndex, length);
                return TextCollection.EffectsLabBlurrinessCheckBoxLabel + checkBoxFeatureLabel;
            }

            var startIndex = ribbonId.IndexOf(TextCollection.DynamicMenuOptionId) + TextCollection.DynamicMenuOptionId.Length;
            var percentage = ribbonId.Substring(startIndex, ribbonId.Length - startIndex);

            return percentage + "% " + TextCollection.EffectsLabBlurrinessTag;
        }
    }
}
