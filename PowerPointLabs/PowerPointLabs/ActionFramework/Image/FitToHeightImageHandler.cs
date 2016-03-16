using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        "fitToHeightShape", 
        "fitToHeightFreeform", 
        "fitToHeightPicture",
        "fitToHeightChart", 
        "fitToHeightTable")]
    class FitToHeightImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.FitToHeight);
        }
    }
}
