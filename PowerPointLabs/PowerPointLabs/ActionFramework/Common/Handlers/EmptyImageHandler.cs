using System.Drawing;

using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Handlers
{
    class EmptyImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return null;
        }
    }
}
