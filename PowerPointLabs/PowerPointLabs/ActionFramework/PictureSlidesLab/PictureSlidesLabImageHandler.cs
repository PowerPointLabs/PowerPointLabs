using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.PictureSlidesLab
{
    [ExportImageRibbonId(TextCollection1.PictureSlidesLabTag)]
    class PictureSlidesLabImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PictureSlidesLab);
        }
    }
}
