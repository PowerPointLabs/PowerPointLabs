using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.TooltipsLab
{
    [ExportLabelRibbonId(TooltipsLabText.ConvertToTooltipTag)]
    class ConvertToTooltipLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TooltipsLabText.ConvertToTooltipButtonLabel;
        }
    }
}
