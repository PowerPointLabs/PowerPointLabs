using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        "FitToWidthMenuShape",
        "FitToWidthMenuFreeform",
        "FitToWidthMenuPicture",
        "FitToWidthMenuGroup",
        "FitToWidthMenuChart",
        "FitToWidthMenuTable")]
    class FitToWidthImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.FitToWidth);
        }
    }
}
