using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        TextCollection.HideSelectedShapeId + TextCollection.MenuShape,
        TextCollection.HideSelectedShapeId + TextCollection.MenuLine,
        TextCollection.HideSelectedShapeId + TextCollection.MenuFreeform,
        TextCollection.HideSelectedShapeId + TextCollection.MenuPicture,
        TextCollection.HideSelectedShapeId + TextCollection.MenuGroup,
        TextCollection.HideSelectedShapeId + TextCollection.MenuInk,
        TextCollection.HideSelectedShapeId + TextCollection.MenuVideo,
        TextCollection.HideSelectedShapeId + TextCollection.MenuTextEdit,
        TextCollection.HideSelectedShapeId + TextCollection.MenuChart,
        TextCollection.HideSelectedShapeId + TextCollection.MenuTable,
        TextCollection.HideSelectedShapeId + TextCollection.MenuTableCell,
        TextCollection.HideSelectedShapeId + TextCollection.MenuSmartArt,
        TextCollection.HideSelectedShapeId + TextCollection.MenuEditSmartArt,
        TextCollection.HideSelectedShapeId + TextCollection.MenuEditSmartArtText)]
    class HideShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.HideSelectedShapeLabel;
        }
    }
}
