using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId("ResizeLabButton")]
    class ResizeLabImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId, string ribbonTag)
        {
            return new Bitmap(Properties.Resources.ResizeLab);
        }
    }
}
