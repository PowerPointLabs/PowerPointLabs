using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SaveLab
{
    [ExportLabelRibbonId(SaveLabText.RibbonMenuId)]
    class SaveLabMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return SaveLabText.RibbonMenuLabel;
        }
    }
}
