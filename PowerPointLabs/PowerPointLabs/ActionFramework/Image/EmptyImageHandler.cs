using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    class EmptyImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId, string ribbonTag)
        {
            return null;
        }
    }
}
