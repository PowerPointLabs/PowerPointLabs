using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportEnabledRibbonId(ShortcutsLabText.AddIntoGroupTag)]
    class AddIntoGroupEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !IsSelectionSingleShape() && !HasPlaceholderInSelection() && !IsSelectionChildShapeRange();
        }
    }
}