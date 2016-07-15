using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.GalleryAction
{
    [ExportGalleryActionRibbonId("EffectsLabBlurSelectedGallery")]
    class EffectsLabBlurrinessGalleryActionHandler : GalleryActionHandler
    {
        protected override void ExecuteGalleryAction(string ribbonId, string selectedId, int selectedIndex)
        {
            this.StartNewUndoEntry();
            
            var selection = this.GetCurrentSelection();
            var slide = this.GetCurrentSlide();

            var feature = ribbonId.Replace("Gallery", "");
            var percentage = int.Parse(selectedId.Replace(feature, ""));

            switch (feature)
            {
                case "EffectsLabBlurSelected":
                    Common.Log.Logger.Log("Entering Effects Lab Blur Selected");
                    EffectsLab.EffectsLabBlurSelected.BlurSelected(slide, selection, percentage);
                    break;
                default:
                    throw new System.Exception("Invalid feature");
            }
        }
    }
}
