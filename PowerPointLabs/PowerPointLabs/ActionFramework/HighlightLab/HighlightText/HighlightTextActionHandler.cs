﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.HighlightLab;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportActionRibbonId(TextCollection.HighlightTextTag)]
    class HighlightTextActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            if (this.GetAddIn().Application.ActiveWindow.Selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                HighlightTextFragments.userSelection = HighlightTextFragments.HighlightTextSelection.kShapeSelected;
            }
            else if (this.GetAddIn().Application.ActiveWindow.Selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
            {
                HighlightTextFragments.userSelection = HighlightTextFragments.HighlightTextSelection.kTextSelected;
            }
            else
            {
                HighlightTextFragments.userSelection = HighlightTextFragments.HighlightTextSelection.kNoneSelected;
            }

            HighlightTextFragments.AddHighlightedTextFragments();
        }
    }
}
