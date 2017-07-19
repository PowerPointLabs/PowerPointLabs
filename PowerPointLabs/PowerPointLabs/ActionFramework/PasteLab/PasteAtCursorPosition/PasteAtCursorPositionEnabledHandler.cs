using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportEnabledRibbonId(TextCollection.PasteAtCursorPositionTag)]
    class PasteAtCursorPositionEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !Graphics.IsClipboardEmpty();
        }
    }
}