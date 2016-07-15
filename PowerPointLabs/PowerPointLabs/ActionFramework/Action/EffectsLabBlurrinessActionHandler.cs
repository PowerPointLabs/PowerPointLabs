using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("EffectsLabBlurSelectedButton")]
    class EffectsLabBlurrinessActionHandler : ActionHandler
    {
        private string feature;
        private Microsoft.Office.Interop.PowerPoint.Selection selection;
        private Models.PowerPointSlide slide;

        protected override void ExecuteAction(string ribbonId)
        {
            feature = ribbonId.Replace("Button", "");
            selection = this.GetCurrentSelection();
            slide = this.GetCurrentSlide();

            var dialog = new Views.EffectsLabBlurrinessDialogBox();
            dialog.SettingsHandler += PropertiesEdited;
            dialog.ShowDialog();
        }

        private void PropertiesEdited(int percentage, bool hasOverlay)
        {
            EffectsLab.EffectsLabBlurSelected.HasOverlay = hasOverlay;

            switch (feature)
            {
                case "EffectsLabBlurSelected":
                    EffectsLab.EffectsLabBlurSelected.BlurSelected(slide, selection, percentage);
                    break;
                default:
                    throw new System.Exception("Invalid feature");
            }
        }
    }
}
