using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportLabelRibbonId(CropLabText.CropToAspectRatioTag + CommonText.RibbonMenu)]
    class CropToAspectRatioOptionsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            int labelStartIndex = 0;
            string label = string.Empty;

            if (ribbonId.Contains(CommonText.DynamicMenuButtonId))
            {
                labelStartIndex = ribbonId.LastIndexOf(CommonText.DynamicMenuButtonId) +
                                  CommonText.DynamicMenuOptionId.Length;
                label = ribbonId.Substring(labelStartIndex);
            }
            else if (ribbonId.Contains(CommonText.DynamicMenuOptionId))
            {
                labelStartIndex = ribbonId.LastIndexOf(CommonText.DynamicMenuOptionId) +
                                  CommonText.DynamicMenuOptionId.Length;
                label = ribbonId.Substring(labelStartIndex).Replace('_', ':');
            }

            return label;
        }
    }
}
