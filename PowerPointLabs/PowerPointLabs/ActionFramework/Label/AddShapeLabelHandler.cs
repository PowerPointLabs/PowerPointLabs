using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId("AddCustomShape", "AddCustomShapePicture", "AddCustomShapeChart", 
                        "AddCustomShapeTable", "AddCustomShapeGroup", "AddCustomShapeFreeform",
                        "AddCustomShapeInk", "AddCustomShapeSmartArt")]
    class AddShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AddCustomShapeShapeLabel;
        }
    }
}
