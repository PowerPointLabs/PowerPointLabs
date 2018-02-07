using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportEnabledRibbonId(ShortcutsLabText.FillSlideTag)]
    class FillSlideEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            //Gets the current selection
            Microsoft.Office.Interop.PowerPoint.Selection currentSelection = this.GetCurrentSelection();
            //Checks if everything currently selected are either shapes or pictures and if number of objects selected equals one
            return ShapeUtil.IsAllPictureOrShape(currentSelection.ShapeRange) && currentSelection.ShapeRange.Count == 1;
        }
    }
}
