using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportEnabledRibbonId(PasteLabText.PasteIntoGroupTag)]
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