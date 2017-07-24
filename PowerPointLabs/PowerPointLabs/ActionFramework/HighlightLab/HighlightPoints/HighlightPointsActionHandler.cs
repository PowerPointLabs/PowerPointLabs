using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.HighlightLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportActionRibbonId(HighlightLabText.HighlightPointsTag)]
    class HighlightPointsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            if (this.GetAddIn().Application.ActiveWindow.Selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kShapeSelected;
            }
            else if (this.GetAddIn().Application.ActiveWindow.Selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
            {
                HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kTextSelected;
            }
            else
            {
                HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kNoneSelected;
            }

            HighlightBulletsText.AddHighlightBulletsText();
        }
    }
}
