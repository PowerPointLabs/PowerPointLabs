using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportLabelRibbonId(HelpText.RibbonMenuId)]
    class HelpMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return HelpText.HelpMenuLabel;
        }
    }
}
