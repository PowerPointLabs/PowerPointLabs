using System.Linq;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.NarrationsLab;
using PowerPointLabs.NarrationsLab.Views;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportActionRibbonId(NarrationsLabText.AddNarrationsTag)]
    class AddNarrationsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            //TODO: This needs to improved to stop using global variables
            this.StartNewUndoEntry();

            var currentSlide = this.GetCurrentSlide();

            if (this.GetCurrentPresentation().SelectedSlides.Any(slide => slide.NotesPageText.Trim() != ""))
            {
                this.GetRibbonUi().RemoveAudioEnabled = true;
                this.GetRibbonUi().RefreshRibbonControl("RemoveNarrationsButton");
            }

            var allAudioFiles = NotesToAudio.EmbedSelectedSlideNotes();

            var recorderPane = this.GetAddIn().GetActivePane(typeof(RecorderTaskPane));

            if (recorderPane == null)
            {
                return;
            }

            var recorder = recorderPane.Control as RecorderTaskPane;

            if (recorder == null)
            {
                return;
            }

            // initialize selected slides' audio
            recorder.InitializeAudioAndScript(this.GetCurrentPresentation().SelectedSlides.ToList(),
                                                  allAudioFiles, true);

            // if current list is visible, update the pane immediately
            if (recorderPane.Visible)
            {
                recorder.UpdateLists(currentSlide.ID);
            }

            if (NarrationsLabSettings.IsPreviewEnabled)
            {
                NotesToAudio.PreviewAnimations();
            }
        }
    }
}
