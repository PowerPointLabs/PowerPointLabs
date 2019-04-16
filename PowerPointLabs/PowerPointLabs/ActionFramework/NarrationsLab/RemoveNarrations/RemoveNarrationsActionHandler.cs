using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.ELearningLab.Views;
using PowerPointLabs.Models;
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

            ComputerVoiceRuntimeService.RemoveAudioFromSelectedSlides();

            CustomTaskPane recorderPane = this.GetAddIn().GetActivePane(typeof(RecorderTaskPane));

            if (recorderPane != null)
            {
                RecorderTaskPane recorder = recorderPane.Control as RecorderTaskPane;
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

            ComputerVoiceRuntimeService.IsRemoveAudioEnabled = false;
            this.GetRibbonUi().RefreshRibbonControl("RemoveNarrationsButton");
        }
    }
}
