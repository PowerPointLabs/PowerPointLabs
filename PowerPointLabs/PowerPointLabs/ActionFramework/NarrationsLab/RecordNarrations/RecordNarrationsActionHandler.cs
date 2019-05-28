using System;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ELearningLab.Views;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportActionRibbonId(NarrationsLabText.RecordNarrationsTag)]
    class RecordNarrationsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            Presentation currentPresentation = this.GetCurrentPresentation().Presentation;
            if (!this.GetRibbonUi().IsValidPresentation(currentPresentation))
            {
                return;
            }

            this.RegisterTaskPane(typeof(RecorderTaskPane), NarrationsLabText.RecManagementPanelTitle,
                null, null);

            CustomTaskPane recorderPane = this.GetAddIn().GetActivePane(typeof(RecorderTaskPane));
            RecorderTaskPane recorder = recorderPane.Control as RecorderTaskPane;

            // if currently the pane is hidden, show the pane
            if (recorder != null && !recorderPane.Visible)
            {
                // fire the pane visble change event
                recorderPane.Visible = true;

                // reload the pane
                recorder.RecorderPaneReload();
            }
        }

        private void TaskPaneVisibleValueChangedEventHandler(object sender, EventArgs e)
        {
            CustomTaskPane recorderPane = Globals.ThisAddIn.GetActivePane(typeof(RecorderTaskPane));

            if (recorderPane == null)
            {
                return;
            }

            RecorderTaskPane recorder = recorderPane.Control as RecorderTaskPane;

            // trigger close form event when closing hide the pane
            if (!recorder?.Visible ?? false)
            {
                recorder.RecorderPaneClosing();
                // remove recorder pane and force it to reload when next time open
                // TODO: Callback to remove task pane from thisaddin, register event.
                Globals.ThisAddIn.RemoveTaskPane(Globals.ThisAddIn.Application.ActiveWindow, typeof(RecorderTaskPane));
            }
        }
    }
}
