using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Enabled.PasteLab
{
    [ExportEnabledRibbonId(
        "PasteAtOriginalPositionMenuShape", "PasteAtOriginalPositionMenuLine", "PasteAtOriginalPositionMenuFreeform",
        "PasteAtOriginalPositionMenuPicture", "PasteAtOriginalPositionMenuGroup", "PasteAtOriginalPositionMenuInk",
        "PasteAtOriginalPositionMenuVideo", "PasteAtOriginalPositionMenuTextEdit", "PasteAtOriginalPositionMenuChart",
        "PasteAtOriginalPositionMenuTable", "PasteAtOriginalPositionMenuTableWhole", "PasteAtOriginalPositionMenuFrame",
        "PasteAtOriginalPositionMenuSmartArtBackground", "PasteAtOriginalPositionMenuSmartArtEditSmartArt",
        "PasteAtOriginalPositionMenuSmartArtEditText", "PasteAtOriginalPositionButton")]
    class PasteAtOriginalPositionEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !Graphics.IsClipboardEmpty();
        }
    }
}