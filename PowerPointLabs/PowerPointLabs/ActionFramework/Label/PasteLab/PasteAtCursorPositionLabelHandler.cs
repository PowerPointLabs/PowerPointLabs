using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuShape,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuLine,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuFreeform,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuPicture,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuGroup,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuInk,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuVideo,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuTextEdit,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuChart,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuTable,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuTableCell,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuSlide,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuSmartArt,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuEditSmartArtText)]
    class PasteAtCursorPositionLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.PasteAtCursorPosition;
        }
    }
}