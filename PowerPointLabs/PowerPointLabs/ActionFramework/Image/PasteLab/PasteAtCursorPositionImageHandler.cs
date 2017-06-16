using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
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
    class PasteAtCursorPositionImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PasteAtCursorPosition);
        }
    }
}