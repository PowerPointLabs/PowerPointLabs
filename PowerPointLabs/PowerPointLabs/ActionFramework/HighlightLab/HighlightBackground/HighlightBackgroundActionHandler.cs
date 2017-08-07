using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.HighlightLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportActionRibbonId(HighlightLabText.HighlightBackgroundTag)]
    class HighlightBackgroundActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            if (this.GetAddIn().Application.ActiveWindow.Selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kShapeSelected;
            }
            else if (this.GetAddIn().Application.ActiveWindow.Selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
            {
                HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kTextSelected;
            }
            else
            {
                HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kNoneSelected;
            }

            HighlightBulletsBackground.AddHighlightBulletsBackground();
        }
    }
}
