using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Enabled.PasteLab
{
    [ExportEnabledRibbonId(
        "PasteAtCursorPositionMenuShape", "PasteAtCursorPositionMenuLine", "PasteAtCursorPositionMenuFreeform",
        "PasteAtCursorPositionMenuPicture", "PasteAtCursorPositionMenuGroup", "PasteAtCursorPositionMenuInk",
        "PasteAtCursorPositionMenuVideo", "PasteAtCursorPositionMenuTextEdit", "PasteAtCursorPositionMenuChart",
        "PasteAtCursorPositionMenuTable", "PasteAtCursorPositionMenuTableWhole", "PasteAtCursorPositionMenuFrame",
        "PasteAtCursorPositionMenuSmartArtBackground", "PasteAtCursorPositionMenuSmartArtEditSmartArt",
        "PasteAtCursorPositionMenuSmartArtEditText")]
    class PasteAtCursorPositionEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !Graphics.IsClipboardEmpty();
        }
    }
}