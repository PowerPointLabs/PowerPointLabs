using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        "AddToShapesLabMenuShape", "AddToShapesLabMenuLine", "AddToShapesLabMenuFreeform",
        "AddToShapesLabMenuPicture", "AddToShapesLabMenuGroup", "AddToShapesLabMenuInk",
        "AddToShapesLabMenuVideo", "AddToShapesLabMenuTextEdit", "AddToShapesLabMenuChart",
        "AddToShapesLabMenuTable", "AddToShapesLabMenuTableWhole", "AddToShapesLabMenuSmartArtBackground",
        "AddToShapesLabMenuSmartArtEditSmartArt", "AddToShapesLabMenuSmartArtEditText")]
    class AddShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AddCustomShapeShapeLabel;
        }
    }
}
