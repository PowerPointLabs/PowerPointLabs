using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        "FitToHeightMenuShape",
        "FitToHeightMenuFreeform",
        "FitToHeightMenuPicture",
        "FitToHeightMenuGroup",
        "FitToHeightMenuChart",
        "FitToHeightMenuTable")]
    class FitToHeightImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.FitToHeight);
        }
    }
}
