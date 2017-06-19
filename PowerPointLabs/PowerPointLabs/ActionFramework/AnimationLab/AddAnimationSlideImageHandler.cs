using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportImageRibbonId("AddAnimationSlide")]
    class AddAnimationSlideImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.AddAnimation);
        }
    }
}
