using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
        "PasteAtCursorPositionMenuFrame",
        "PasteAtCursorPositionMenuShape",
        "PasteAtCursorPositionMenuLine",
        "PasteAtCursorPositionMenuFreeform",
        "PasteAtCursorPositionMenuPicture",
        "PasteAtCursorPositionMenuGroup",
        "PasteAtCursorPositionMenuInk",
        "PasteAtCursorPositionMenuVideo",
        "PasteAtCursorPositionMenuChart",
        "PasteAtCursorPositionMenuTable",
        "PasteAtCursorPositionMenuTableWhole",
        "PasteAtCursorPositionMenuSmartArtBackground",
        "PasteAtCursorPositionMenuSmartArtEditSmartArt",
        "PasteAtCursorPositionMenuSmartArtEditText")]
    class PasteAtCursorPositionImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PasteLab);
        }
    }
}