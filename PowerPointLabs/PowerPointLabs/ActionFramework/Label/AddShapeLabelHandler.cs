using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId("AddCustomShape")]
    class AddShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AddCustomShapeShapeLabel;
        }
    }
}
