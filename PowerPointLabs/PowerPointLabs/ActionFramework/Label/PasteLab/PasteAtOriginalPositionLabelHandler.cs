using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuShape,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuLine,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuFreeform,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuPicture,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuGroup,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuInk,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuVideo,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuTextEdit,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuChart,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuTable,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuTableCell,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuSlide,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuSmartArt,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuEditSmartArtText,
        TextCollection.PasteAtOriginalPositionId + TextCollection.RibbonButton)]
    class PasteAtOriginalPositionLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.PasteAtOriginalPosition;
        }
    }
}