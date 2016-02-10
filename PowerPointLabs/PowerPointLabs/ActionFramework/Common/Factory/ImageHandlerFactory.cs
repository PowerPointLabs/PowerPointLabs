using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Image;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for ImageHandler
    /// </summary>
    public class ImageHandlerFactory : BaseHandlerFactory<ImageHandler>
    {
        protected override ImageHandler GetEmptyHandler()
        {
            return new EmptyImageHandler();
        }
    }
}
