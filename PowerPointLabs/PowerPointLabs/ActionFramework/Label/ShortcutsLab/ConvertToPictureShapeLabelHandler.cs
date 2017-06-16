using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        TextCollection.ConvertToPictureTag + TextCollection.MenuShape,
        TextCollection.ConvertToPictureTag + TextCollection.MenuLine,
        TextCollection.ConvertToPictureTag + TextCollection.MenuFreeform,
        TextCollection.ConvertToPictureTag + TextCollection.MenuGroup,
        TextCollection.ConvertToPictureTag + TextCollection.MenuInk,
        TextCollection.ConvertToPictureTag + TextCollection.MenuVideo,
        TextCollection.ConvertToPictureTag + TextCollection.MenuTextEdit,
        TextCollection.ConvertToPictureTag + TextCollection.MenuChart,
        TextCollection.ConvertToPictureTag + TextCollection.MenuTable,
        TextCollection.ConvertToPictureTag + TextCollection.MenuTableCell,
        TextCollection.ConvertToPictureTag + TextCollection.MenuSmartArt,
        TextCollection.ConvertToPictureTag + TextCollection.MenuEditSmartArt,
        TextCollection.ConvertToPictureTag + TextCollection.MenuEditSmartArtText)]
    class ConvertToPictureShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.ConvertToPictureShapeLabel;
        }
    }
}
