using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportEnabledRibbonId(TextCollection.AddIntoGroupTag)]
    class AddIntoGroupEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !IsSelectionSingleShape() && !HasPlaceholderInSelection() && !IsSelectionChildShapeRange();
        }
    }
}