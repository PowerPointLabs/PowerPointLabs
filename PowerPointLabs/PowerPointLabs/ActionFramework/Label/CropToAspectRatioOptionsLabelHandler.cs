using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(TextCollection.CropToAspectRatioTag)]
    class CropToAspectRatioOptionsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            int labelStartIndex = 0;
            string label = string.Empty;

            if (ribbonId.Contains(TextCollection.DynamicMenuButtonId))
            {
                labelStartIndex = ribbonId.LastIndexOf(TextCollection.DynamicMenuButtonId) +
                                  TextCollection.DynamicMenuOptionId.Length;
                label = ribbonId.Substring(labelStartIndex);
            }
            else if (ribbonId.Contains(TextCollection.DynamicMenuOptionId))
            {
                labelStartIndex = ribbonId.LastIndexOf(TextCollection.DynamicMenuOptionId) +
                                  TextCollection.DynamicMenuOptionId.Length;
                label = ribbonId.Substring(labelStartIndex).Replace('_', ':');
            }

            return label;
        }
    }
}
