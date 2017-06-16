using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuShape,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuLine,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuFreeform,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuPicture,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuGroup,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuInk,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuVideo,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuTextEdit,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuChart,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuTable,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuTableCell,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuSmartArt,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuEditSmartArt,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuEditSmartArtText)]
    class AddShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AddCustomShapeShapeLabel;
        }
    }
}
