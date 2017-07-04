using PowerPointLabs.ActionFramework.Common.Handlers;
using PowerPointLabs.ActionFramework.Common.Interface;

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
