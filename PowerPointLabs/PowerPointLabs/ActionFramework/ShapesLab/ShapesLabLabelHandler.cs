using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapesLab
{
    [ExportLabelRibbonId(TextCollection1.ShapesLabTag)]
    class ShapesLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return ShapesLabText.RibbonMenuLabel;
        }
    }
}
