using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        TextCollection.AddCustomShapeTag + TextCollection.MenuShape,
        TextCollection.AddCustomShapeTag + TextCollection.MenuLine,
        TextCollection.AddCustomShapeTag + TextCollection.MenuFreeform,
        TextCollection.AddCustomShapeTag + TextCollection.MenuPicture,
        TextCollection.AddCustomShapeTag + TextCollection.MenuGroup,
        TextCollection.AddCustomShapeTag + TextCollection.MenuInk,
        TextCollection.AddCustomShapeTag + TextCollection.MenuVideo,
        TextCollection.AddCustomShapeTag + TextCollection.MenuTextEdit,
        TextCollection.AddCustomShapeTag + TextCollection.MenuChart,
        TextCollection.AddCustomShapeTag + TextCollection.MenuTable,
        TextCollection.AddCustomShapeTag + TextCollection.MenuTableCell,
        TextCollection.AddCustomShapeTag + TextCollection.MenuSmartArt,
        TextCollection.AddCustomShapeTag + TextCollection.MenuEditSmartArt,
        TextCollection.AddCustomShapeTag + TextCollection.MenuEditSmartArtText)]
    class AddShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AddCustomShapeShapeLabel;
        }
    }
}
