using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CaptionsLab
{
    [ExportActionRibbonId(CaptionsLabText.RemoveNotesTag)]
    class RemoveNotesActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            //TODO: This needs to improved to stop using global variables
            this.StartNewUndoEntry();

            foreach (var slide in this.GetCurrentPresentation().SelectedSlides)
            {
                slide.NotesPageText = string.Empty;
            }
        }
    }
}
