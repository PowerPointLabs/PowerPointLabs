using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(TextCollection.EffectsLabBlurrinessTag)]
    class EffectsLabBlurrinessLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId, string ribbonTag)
        {
            if (ribbonId.Contains("Button"))
            {
                return TextCollection.EffectsLabBlurrinessButtonLabel;
            }

            if (ribbonId.Contains("CheckBox"))
            {
                return TextCollection.EffectsLabBlurrinessCheckBoxLabel;
            }

            var startIndex = ribbonId.IndexOf("Option") + 6;
            var percentage = ribbonId.Substring(startIndex, ribbonId.Length - startIndex);

            return percentage + "% " + ribbonTag;
        }
    }
}
