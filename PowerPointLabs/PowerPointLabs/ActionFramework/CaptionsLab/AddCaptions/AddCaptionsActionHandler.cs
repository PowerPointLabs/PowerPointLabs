using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.CaptionsLab;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CaptionsLab
{
    [ExportActionRibbonId(CaptionsLabText.AddCaptionsTag)]
    class AddCaptionsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            //TODO: This needs to improved to stop using global variables
            this.StartNewUndoEntry();

            foreach (PowerPointSlide slide in this.GetCurrentPresentation().SelectedSlides)
            {
                if (slide.NotesPageText.Trim() != "")
                {
                    this.GetRibbonUi().RemoveCaptionsEnabled = true;
                    break;
                }
            }

            NotesToCaptions.EmbedCaptionsOnSelectedSlides();
            this.GetRibbonUi().RefreshRibbonControl("RemoveCaptionsButton");
        }
    }
}
