using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Enabled.PasteLab
{
    [ExportEnabledRibbonId(
        "PasteToFillSlideMenuShape", "PasteToFillSlideMenuLine", "PasteToFillSlideMenuFreeform",
        "PasteToFillSlideMenuPicture", "PasteToFillSlideMenuGroup", "PasteToFillSlideMenuInk",
        "PasteToFillSlideMenuVideo", "PasteToFillSlideMenuTextEdit", "PasteToFillSlideMenuChart",
        "PasteToFillSlideMenuTable", "PasteToFillSlideMenuTableWhole", "PasteToFillSlideMenuFrame",
        "PasteToFillSlideMenuSmartArtBackground", "PasteToFillSlideMenuSmartArtEditSmartArt",
        "PasteToFillSlideMenuSmartArtEditText", "PasteToFillSlideButton")]
    class PasteToFillSlideEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !Graphics.IsClipboardEmpty();
        }
    }
}
