using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportLabelRibbonId(TextCollection1.CropToAspectRatioTag + TextCollection1.RibbonMenu)]
    class CropToAspectRatioOptionsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            int labelStartIndex = 0;
            string label = string.Empty;

            if (ribbonId.Contains(TextCollection1.DynamicMenuButtonId))
            {
                labelStartIndex = ribbonId.LastIndexOf(TextCollection1.DynamicMenuButtonId) +
                                  TextCollection1.DynamicMenuOptionId.Length;
                label = ribbonId.Substring(labelStartIndex);
            }
            else if (ribbonId.Contains(TextCollection1.DynamicMenuOptionId))
            {
                labelStartIndex = ribbonId.LastIndexOf(TextCollection1.DynamicMenuOptionId) +
                                  TextCollection1.DynamicMenuOptionId.Length;
                label = ribbonId.Substring(labelStartIndex).Replace('_', ':');
            }

            return label;
        }
    }
}
