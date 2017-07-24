using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Label.HighlightLab
{
    [ExportLabelRibbonId(HighlightLabText.RibbonMenuId)]
    class HighlightLabMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return HighlightLabText.RibbonMenuLabel;
        }
    }
}
