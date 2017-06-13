using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
        "PasteAtOriginalPositionMenuFrame",
        "PasteAtOriginalPositionMenuShape",
        "PasteAtOriginalPositionMenuLine",
        "PasteAtOriginalPositionMenuFreeform",
        "PasteAtOriginalPositionMenuPicture",
        "PasteAtOriginalPositionMenuGroup",
        "PasteAtOriginalPositionMenuInk",
        "PasteAtOriginalPositionMenuVideo",
        "PasteAtOriginalPositionMenuTextEdit",
        "PasteAtOriginalPositionMenuChart",
        "PasteAtOriginalPositionMenuTable",
        "PasteAtOriginalPositionMenuTableWhole",
        "PasteAtOriginalPositionMenuSmartArtBackground",
        "PasteAtOriginalPositionMenuSmartArtEditSmartArt",
        "PasteAtOriginalPositionMenuSmartArtEditText")]
    class PasteAtOriginalPositionImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PasteLab);
        }
    }
}