using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        TextCollection.ConvertToPictureId + TextCollection.MenuShape,
        TextCollection.ConvertToPictureId + TextCollection.MenuLine,
        TextCollection.ConvertToPictureId + TextCollection.MenuFreeform,
        TextCollection.ConvertToPictureId + TextCollection.MenuGroup,
        TextCollection.ConvertToPictureId + TextCollection.MenuInk,
        TextCollection.ConvertToPictureId + TextCollection.MenuVideo,
        TextCollection.ConvertToPictureId + TextCollection.MenuTextEdit,
        TextCollection.ConvertToPictureId + TextCollection.MenuChart,
        TextCollection.ConvertToPictureId + TextCollection.MenuTable,
        TextCollection.ConvertToPictureId + TextCollection.MenuTableCell,
        TextCollection.ConvertToPictureId + TextCollection.MenuSmartArt,
        TextCollection.ConvertToPictureId + TextCollection.MenuEditSmartArt,
        TextCollection.ConvertToPictureId + TextCollection.MenuEditSmartArtText)]
    class ConvertToPictureShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.ConvertToPictureShapeLabel;
        }
    }
}
