using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        "fitToWidthShape",
        "fitToWidthFreeform",
        "fitToWidthPicture",
        "fitToWidthChart",
        "fitToWidthTable")]
    class FitToWidthImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId, string ribbonTag)
        {
            return new Bitmap(Properties.Resources.FitToWidth);
        }
    }
}
