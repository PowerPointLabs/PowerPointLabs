using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.CaptionsLab
{
    [ExportActionRibbonId(TextCollection1.RemoveNotesTag)]
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
