using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.CaptionsLab;

namespace PowerPointLabs.ActionFramework.CaptionsLab
{
    [ExportActionRibbonId(TextCollection1.RemoveCaptionsTag)]
    class RemoveCaptionsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            //TODO: This needs to improved to stop using global variables
            this.StartNewUndoEntry();

            this.GetRibbonUi().RemoveCaptionsEnabled = false;
            this.GetRibbonUi().RefreshRibbonControl("RemoveCaptionsButton");
            NotesToCaptions.RemoveCaptionsFromSelectedSlides();
        }
    }
}
