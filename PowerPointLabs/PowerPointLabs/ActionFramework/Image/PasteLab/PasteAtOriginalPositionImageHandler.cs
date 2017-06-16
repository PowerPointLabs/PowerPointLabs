using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuShape,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuLine,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuFreeform,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuPicture,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuGroup,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuInk,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuVideo,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuTextEdit,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuChart,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuTable,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuTableCell,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuSlide,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuSmartArt,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuEditSmartArtText,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.RibbonButton)]
    class PasteAtOriginalPositionImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PasteAtOriginalPosition);
        }
    }
}