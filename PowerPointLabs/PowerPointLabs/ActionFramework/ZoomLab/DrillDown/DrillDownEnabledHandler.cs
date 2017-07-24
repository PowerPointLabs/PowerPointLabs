using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportEnabledRibbonId(ZoomLabText.DrillDownTag)]
    class DrillDownEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            Selection currentSelection = this.GetCurrentSelection();
            return ShapeUtil.IsSelectionSingleShape(currentSelection) && 
                ShapeUtil.IsSelectionAllRectangle(currentSelection);
        }
    }
}