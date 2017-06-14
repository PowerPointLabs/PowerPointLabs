using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
        "PasteAtOriginalPositionMenuShape", "PasteAtOriginalPositionMenuLine", "PasteAtOriginalPositionMenuFreeform",
        "PasteAtOriginalPositionMenuPicture", "PasteAtOriginalPositionMenuGroup", "PasteAtOriginalPositionMenuInk",
        "PasteAtOriginalPositionMenuVideo", "PasteAtOriginalPositionMenuTextEdit", "PasteAtOriginalPositionMenuChart",
        "PasteAtOriginalPositionMenuTable", "PasteAtOriginalPositionMenuTableWhole", "PasteAtOriginalPositionMenuFrame",
        "PasteAtOriginalPositionMenuSmartArtBackground", "PasteAtOriginalPositionMenuSmartArtEditSmartArt",
        "PasteAtOriginalPositionMenuSmartArtEditText", "PasteAtOriginalPositionButton")]
    class PasteAtOriginalPositionImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PasteAtOriginalPosition);
        }
    }
}