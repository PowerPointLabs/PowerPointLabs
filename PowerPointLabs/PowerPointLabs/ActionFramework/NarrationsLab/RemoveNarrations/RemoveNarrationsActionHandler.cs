using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.NarrationsLab;
using PowerPointLabs.NarrationsLab.Views;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportActionRibbonId(NarrationsLabText.RemoveNarrationsTag)]
    class RemoveNarrationsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            //TODO: This needs to improved to stop using global variables (Change RemoveAudioEnabledHandler too)
            this.StartNewUndoEntry();

            NotesToAudio.RemoveAudioFromSelectedSlides();

            var recorderPane = this.GetAddIn().GetActivePane(typeof(RecorderTaskPane));

            if (recorderPane != null)
            {
                var recorder = recorderPane.Control as RecorderTaskPane;
                recorder.ClearRecordDataListForSelectedSlides();

                // if current list is visible, update the pane immediately
                if (recorderPane.Visible)
                {
                    foreach (PowerPointSlide slide in this.GetCurrentPresentation().SelectedSlides)
                    {
                        recorder.UpdateLists(slide.ID);
                    }
                }
            }

            this.GetRibbonUi().RemoveAudioEnabled = false;
            this.GetRibbonUi().RefreshRibbonControl("RemoveNarrationsButton");
        }
    }
}
