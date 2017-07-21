using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportEnabledRibbonId(TextCollection.AddIntoGroupTag)]
    class AddIntoGroupEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            Selection currentSelection = this.GetCurrentSelection();
            return !ShapeUtil.IsSelectionSingleShape(currentSelection) && 
                !ShapeUtil.HasPlaceholderInSelection(currentSelection) && 
                !ShapeUtil.IsSelectionChildShapeRange(currentSelection);
        }
    }
}