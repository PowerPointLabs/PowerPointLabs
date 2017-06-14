using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        "HideShapeMenuShape", "HideShapeMenuLine", "HideShapeMenuFreeform",
        "HideShapeMenuPicture", "HideShapeMenuGroup", "HideShapeMenuInk",
        "HideShapeMenuVideo", "HideShapeMenuTextEdit", "HideShapeMenuChart",
        "HideShapeMenuTable", "HideShapeMenuTableWhole", "HideShapeMenuSmartArtBackground",
        "HideShapeMenuSmartArtEditSmartArt", "HideShapeMenuSmartArtEditText")]
    class HideShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.HideSelectedShapeLabel;
        }
    }
}
