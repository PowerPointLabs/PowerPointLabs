using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapesLab.ShapesLabMenu
{
    [ExportLabelRibbonId(ShapesLabText.RibbonMenuId)]
    class ShapesLabMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return ShapesLabText.RibbonMenuLabel;
        }
    }
}