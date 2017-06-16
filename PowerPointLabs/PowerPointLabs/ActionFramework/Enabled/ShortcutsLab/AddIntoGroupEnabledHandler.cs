using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Enabled.ShortcutsLab
{
    [ExportEnabledRibbonId(TextCollection.AddIntoGroupId + TextCollection.MenuGroup)]
    class AddIntoGroupEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !IsSelectionSingleShape() && !HasPlaceholderInSelection() && !IsSelectionChildShapeRange();
        }
    }
}