using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        "ConvertToPictureMenuShape", "ConvertToPictureMenuLine", "ConvertToPictureMenuFreeform",
        "ConvertToPictureMenuGroup", "ConvertToPictureMenuInk", "ConvertToPictureMenuVideo",
        "ConvertToPictureMenuTextEdit", "ConvertToPictureMenuChart", "ConvertToPictureMenuTable",
        "ConvertToPictureMenuTableWhole", "ConvertToPictureMenuSmartArtBackground", "ConvertToPictureMenuSmartArtEditSmartArt",
        "ConvertToPictureMenuSmartArtEditText")]
    class ConvertToPictureImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.ConvertToPicture);
        }
    }
}
