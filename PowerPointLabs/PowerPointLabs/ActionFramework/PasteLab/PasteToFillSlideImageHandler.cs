using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportImageRibbonId(TextCollection.PasteToFillSlideTag)]
    class PasteToFillSlideImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PasteToFillSlide);
        }
    }
}
