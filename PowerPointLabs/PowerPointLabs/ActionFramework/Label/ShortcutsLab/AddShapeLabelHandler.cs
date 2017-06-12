using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId("AddToShapesLabMenuShape", "AddToShapesLabMenuLine", "AddToShapesLabMenuPicture",
                        "AddToShapesLabMenuChart", "AddToShapesLabMenuTable", "AddToShapesLabMenuGroup",
                        "AddToShapesLabMenuFreeform", "AddToShapesLabMenuInk", "AddToShapesLabMenuSmartArt")]
    class AddShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AddCustomShapeShapeLabel;
        }
    }
}
