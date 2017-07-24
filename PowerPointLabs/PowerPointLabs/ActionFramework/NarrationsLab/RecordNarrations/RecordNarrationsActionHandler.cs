using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.NarrationsLab.Views;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportActionRibbonId(NarrationsLabText.RecordNarrationsTag)]
    class RecordNarrationsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            //TODO: This needs to improved to stop using global variables
            var currentPresentation = this.GetCurrentPresentation().Presentation;

            if (!this.GetRibbonUi().IsValidPresentation(currentPresentation))
            {
                return;
            }

            // prepare media files
            var tempPath = this.GetAddIn().PrepareTempFolder(currentPresentation);
            this.GetAddIn().PrepareMediaFiles(currentPresentation, tempPath);

            this.GetAddIn().RegisterRecorderPane(currentPresentation.Windows[1], tempPath);

            var recorderPane = this.GetAddIn().GetActivePane(typeof(RecorderTaskPane));
            var recorder = recorderPane.Control as RecorderTaskPane;

            // if currently the pane is hidden, show the pane
            if (recorder != null && !recorderPane.Visible)
            {
                // fire the pane visble change event
                recorderPane.Visible = true;

                // reload the pane
                recorder.RecorderPaneReload();
            }
        }
    }
}
