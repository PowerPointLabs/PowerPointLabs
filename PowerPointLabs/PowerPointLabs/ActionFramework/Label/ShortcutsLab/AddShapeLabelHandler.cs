using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId("AddToShapesLabMenuShape", "AddToShapesLabMenuLine", "AddToShapesLabMenuPicture",
                        "AddToShapesLabMenuChart", "AddToShapesLabMenuTable", "AddToShapesLabMenuGroup", "AddToShapesLabMenuVideo",
                        "AddToShapesLabMenuTextEdit", "AddToShapesLabMenuFreeform", "AddToShapesLabMenuInk", "AddToShapesLabMenuSmartArtBackground",
                        "AddToShapesLabMenuTableWhole", "AddToShapesLabMenuSmartArtEditSmartArt", "AddToShapesLabMenuSmartArtEditText")]
    class AddShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AddCustomShapeShapeLabel;
        }
    }
}
