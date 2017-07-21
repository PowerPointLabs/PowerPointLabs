using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Label.HighlightLab
{
    [ExportLabelRibbonId(TextCollection1.HighlightLabMenuId)]
    class HighlightLabMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return HighlightLabText.RibbonMenuLabel;
        }
    }
}
