using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.GalleryAction;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for GalleryActionHandler
    /// </summary>
    public class GalleryActionHandlerFactory : BaseHandlerFactory<GalleryActionHandler>
    {
        protected override GalleryActionHandler GetEmptyHandler()
        {
            return new EmptyGalleryActionHandler();
        }
    }
}
