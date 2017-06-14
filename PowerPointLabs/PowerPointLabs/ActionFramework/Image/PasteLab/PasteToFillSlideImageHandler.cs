using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
        "PasteToFillSlideMenuShape", "PasteToFillSlideMenuLine", "PasteToFillSlideMenuFreeform",
        "PasteToFillSlideMenuPicture", "PasteToFillSlideMenuGroup", "PasteToFillSlideMenuInk",
        "PasteToFillSlideMenuVideo", "PasteToFillSlideMenuTextEdit", "PasteToFillSlideMenuChart",
        "PasteToFillSlideMenuTable", "PasteToFillSlideMenuTableWhole", "PasteToFillSlideMenuFrame",
        "PasteToFillSlideMenuSmartArtBackground", "PasteToFillSlideMenuSmartArtEditSmartArt",
        "PasteToFillSlideMenuSmartArtEditText", "PasteToFillSlideButton")]
    class PasteToFillSlideImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PasteToFillSlide);
        }
    }
}
