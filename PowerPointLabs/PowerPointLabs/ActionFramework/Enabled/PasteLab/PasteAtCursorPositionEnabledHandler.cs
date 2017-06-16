using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Enabled.PasteLab
{
    [ExportEnabledRibbonId(
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
    class PasteAtCursorPositionEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !Graphics.IsClipboardEmpty();
        }
    }
}