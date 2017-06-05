using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.Group
{
    [ExportImageRibbonId("AutoMotionGroup")]
    class AnimationLabImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.AnimationLabSmall);
        }
    }
}
