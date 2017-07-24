using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportLabelRibbonId(PasteLabText.RibbonMenuId)]
    class PasteLabMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return PasteLabText.RibbonMenuLabel;
        }
    }
}
