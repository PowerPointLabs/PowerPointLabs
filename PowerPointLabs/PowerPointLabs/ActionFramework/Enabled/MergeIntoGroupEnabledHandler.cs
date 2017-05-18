using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Enabled
{
    [ExportEnabledRibbonId("MergeIntoGroup")]
    class MergeIntoGroupEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !HasPlaceholderInSelection();
        }
    }
}