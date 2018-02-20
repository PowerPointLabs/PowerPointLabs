using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.SaveLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SaveLab
{
    [ExportActionRibbonId(SaveLabText.SavePresentationsButtonTag)]
    class SaveLabButtonActionHandler : Common.Interface.ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            // Save action here

            // Get selection of the slides
            Selection currentSelection = this.GetCurrentSelection();
             //check the type of selection and ensure that it is a slide and that the number of slides selected is >= 1
            if (currentSelection.Type == PpSelectionType.ppSelectionSlides && currentSelection.SlideRange.Count >= 1)
            {
                // Perform the actual save action
                SaveLabMain.SaveFile(currentSelection.SlideRange);
            }
            else
            {
                //if no slides return error message or do nothing
                return;
            }
            // Save selection as a separate ppt file
        }
    }
}
