using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
        "PasteAtCursorPositionMenuShape", "PasteAtCursorPositionMenuLine", "PasteAtCursorPositionMenuFreeform",
        "PasteAtCursorPositionMenuPicture", "PasteAtCursorPositionMenuGroup", "PasteAtCursorPositionMenuInk",
        "PasteAtCursorPositionMenuVideo", "PasteAtCursorPositionMenuTextEdit", "PasteAtCursorPositionMenuChart",
        "PasteAtCursorPositionMenuTable", "PasteAtCursorPositionMenuTableWhole", "PasteAtCursorPositionMenuFrame",
        "PasteAtCursorPositionMenuSmartArtBackground", "PasteAtCursorPositionMenuSmartArtEditSmartArt",
        "PasteAtCursorPositionMenuSmartArtEditText")]
    class PasteAtCursorPositionImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PasteAtCursorPosition);
        }
    }
}