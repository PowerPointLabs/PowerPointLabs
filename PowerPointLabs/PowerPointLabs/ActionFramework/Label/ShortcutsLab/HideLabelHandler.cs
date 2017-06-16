using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        TextCollection.HideSelectedShapeTag + TextCollection.MenuShape,
        TextCollection.HideSelectedShapeTag + TextCollection.MenuLine,
        TextCollection.HideSelectedShapeTag + TextCollection.MenuFreeform,
        TextCollection.HideSelectedShapeTag + TextCollection.MenuPicture,
        TextCollection.HideSelectedShapeTag + TextCollection.MenuGroup,
        TextCollection.HideSelectedShapeTag + TextCollection.MenuInk,
        TextCollection.HideSelectedShapeTag + TextCollection.MenuVideo,
        TextCollection.HideSelectedShapeTag + TextCollection.MenuTextEdit,
        TextCollection.HideSelectedShapeTag + TextCollection.MenuChart,
        TextCollection.HideSelectedShapeTag + TextCollection.MenuTable,
        TextCollection.HideSelectedShapeTag + TextCollection.MenuTableCell,
        TextCollection.HideSelectedShapeTag + TextCollection.MenuSmartArt,
        TextCollection.HideSelectedShapeTag + TextCollection.MenuEditSmartArt,
        TextCollection.HideSelectedShapeTag + TextCollection.MenuEditSmartArtText)]
    class HideShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.HideSelectedShapeLabel;
        }
    }
}
