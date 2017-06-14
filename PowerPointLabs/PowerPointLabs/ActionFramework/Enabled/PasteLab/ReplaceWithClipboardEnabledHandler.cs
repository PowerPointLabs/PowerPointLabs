using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Enabled.PasteLab
{
    [ExportEnabledRibbonId(
        "ReplaceWithClipboardMenuShape", "ReplaceWithClipboardMenuLine", "ReplaceWithClipboardMenuFreeform",
        "ReplaceWithClipboardMenuPicture", "ReplaceWithClipboardMenuGroup", "ReplaceWithClipboardMenuInk",
        "ReplaceWithClipboardMenuVideo", "ReplaceWithClipboardMenuTextEdit", "ReplaceWithClipboardMenuChart",
        "ReplaceWithClipboardMenuTable", "ReplaceWithClipboardMenuTableWhole", "ReplaceWithClipboardMenuSmartArtBackground",
        "ReplaceWithClipboardButton")]
    class ReplaceWithClipboardEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !Graphics.IsClipboardEmpty() && IsSelectionSingleShape();
        }
    }
}