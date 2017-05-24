using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Enabled.PasteLab
{
    [ExportEnabledRibbonId(
        "PasteAtCursorPosition",
        "PasteAtCursorPositionShape",
        "PasteAtCursorPositionFreeform",
        "PasteAtCursorPositionPicture",
        "PasteAtCursorPositionGroup")]
    class PasteAtCursorPositionEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !Graphics.IsClipboardEmpty();
        }
    }
}