using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuShape,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuLine,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuFreeform,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuPicture,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuGroup,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuInk,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuVideo,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuTextEdit,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuChart,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuTable,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuTableCell,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuSlide,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuSmartArt,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteToFillSlideMenuId + TextCollection.MenuEditSmartArtText,
        TextCollection.PasteToFillSlideMenuId + TextCollection.RibbonButton)]
    class PasteToFillSlideImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PasteToFillSlide);
        }
    }
}
