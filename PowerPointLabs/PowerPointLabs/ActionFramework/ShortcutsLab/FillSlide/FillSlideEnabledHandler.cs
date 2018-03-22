using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportEnabledRibbonId(ShortcutsLabText.FillSlideTag)]
    class FillSlideEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            Microsoft.Office.Interop.PowerPoint.Selection currentSelection = this.GetCurrentSelection();

            return ShapeUtil.IsAllPictureOrShape(currentSelection.ShapeRange) && currentSelection.ShapeRange.Count == 1;
        }
    }
}
