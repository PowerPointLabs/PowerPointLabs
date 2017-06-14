using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.HelpMenu
{
    [ExportLabelRibbonId("HelpMenu")]
    class HelpMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.HelpMenuLabel;
        }
    }
}
