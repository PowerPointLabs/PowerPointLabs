using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportEnabledRibbonId(EffectsLabText.AddSpotlightTag)]
    class AddSpotlightEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            Selection currentSelection = this.GetCurrentSelection();
            return ShapeUtil.IsSelectionAllShapeWithArea(currentSelection);
        }
    }
}