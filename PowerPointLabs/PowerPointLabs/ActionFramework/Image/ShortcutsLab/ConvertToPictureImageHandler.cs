using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        TextCollection.ConvertToPictureId + TextCollection.MenuShape,
        TextCollection.ConvertToPictureId + TextCollection.MenuLine,
        TextCollection.ConvertToPictureId + TextCollection.MenuFreeform,
        TextCollection.ConvertToPictureId + TextCollection.MenuGroup,
        TextCollection.ConvertToPictureId + TextCollection.MenuInk,
        TextCollection.ConvertToPictureId + TextCollection.MenuVideo,
        TextCollection.ConvertToPictureId + TextCollection.MenuTextEdit,
        TextCollection.ConvertToPictureId + TextCollection.MenuChart,
        TextCollection.ConvertToPictureId + TextCollection.MenuTable,
        TextCollection.ConvertToPictureId + TextCollection.MenuTableCell,
        TextCollection.ConvertToPictureId + TextCollection.MenuSmartArt,
        TextCollection.ConvertToPictureId + TextCollection.MenuEditSmartArt,
        TextCollection.ConvertToPictureId + TextCollection.MenuEditSmartArtText)]
    class ConvertToPictureImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.ConvertToPicture);
        }
    }
}
