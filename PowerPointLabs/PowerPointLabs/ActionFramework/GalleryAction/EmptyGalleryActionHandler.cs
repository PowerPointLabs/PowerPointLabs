using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.GalleryAction
{
    /// <summary>
    /// the gallery action handler that does nothing
    /// </summary>
    class EmptyGalleryActionHandler : GalleryActionHandler
    {
        protected override void ExecuteGalleryAction(string ribbonId, string selectedId, int selectedIndex)
        {
        }
    }
}
