using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.SaveLab;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.SaveLab
{
    [ExportActionRibbonId(SaveLabText.SavePresentationsButtonTag)]
    class SaveLabButtonActionHandler : Common.Interface.ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            Selection currentSelection = this.GetCurrentSelection();
            Models.PowerPointPresentation currentPresentation = this.GetCurrentPresentation();
             // Check the type of selection and ensure that it is a slide and that the number of slides selected is >= 1
            if (currentSelection.Type == PpSelectionType.ppSelectionSlides && currentSelection.SlideRange.Count >= 1)
            {
                // Perform the actual save action
                SaveLabMain.SaveFile(currentPresentation);
            }
            else
            {
                // If no slides return error message or do nothing
                MessageBoxUtil.Show(SaveLabText.ErrorZeroSlidesSelected, CommonText.ErrorSlideSelectionTitle);
            }
        }
    }
}
