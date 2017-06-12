using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        "HideShapeMenuShape",
        "HideShapeMenuLine",
        "HideShapeMenuFreeform",
        "HideShapeMenuPicture",
        "HideShapeMenuGroup",
        "HideShapeMenuChart",
        "HideShapeMenuTable")]
    class HideImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.HideShape);
        }
    }
}
