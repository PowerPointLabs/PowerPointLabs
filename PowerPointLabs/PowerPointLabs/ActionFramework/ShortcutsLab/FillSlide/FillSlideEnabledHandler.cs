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
            //Checks if the current selection is either a shape or picture and enable if true
            return ShapeUtil.IsAllPictureOrShape(currentSelection.ShapeRange);
        }
    }
}
