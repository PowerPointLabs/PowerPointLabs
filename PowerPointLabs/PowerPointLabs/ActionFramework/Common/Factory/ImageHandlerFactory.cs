using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Image;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    public class ImageHandlerFactory : BaseHandlerFactory<ImageHandler>
    {
        protected override ImageHandler GetEmptyHandler()
        {
            return new EmptyImageHandler();
        }
    }
}
