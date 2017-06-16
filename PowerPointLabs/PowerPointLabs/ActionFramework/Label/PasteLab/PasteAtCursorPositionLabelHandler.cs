using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuShape,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuLine,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuFreeform,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuPicture,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuGroup,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuInk,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuVideo,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuTextEdit,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuChart,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuTable,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuTableCell,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuSlide,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuSmartArt,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuEditSmartArtText)]
    class PasteAtCursorPositionLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.PasteAtCursorPosition;
        }
    }
}