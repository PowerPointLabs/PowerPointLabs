using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Enabled.PasteLab
{
    [ExportEnabledRibbonId(
        TextCollection.PasteIntoGroupMenuId + TextCollection.MenuGroup,
        TextCollection.PasteIntoGroupMenuId + TextCollection.RibbonButton)]
    class PasteIntoGroupEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !Graphics.IsClipboardEmpty() && 
                IsSelectionMultipleOrGroup() &&
                !HasPlaceholderInSelection();
        }
    }
}