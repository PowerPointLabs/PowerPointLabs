using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Enabled.PasteLab
{
    [ExportEnabledRibbonId(
        TextCollection.PasteToFillSlideId + TextCollection.MenuShape,
        TextCollection.PasteToFillSlideId + TextCollection.MenuLine,
        TextCollection.PasteToFillSlideId + TextCollection.MenuFreeform,
        TextCollection.PasteToFillSlideId + TextCollection.MenuPicture,
        TextCollection.PasteToFillSlideId + TextCollection.MenuGroup,
        TextCollection.PasteToFillSlideId + TextCollection.MenuInk,
        TextCollection.PasteToFillSlideId + TextCollection.MenuVideo,
        TextCollection.PasteToFillSlideId + TextCollection.MenuTextEdit,
        TextCollection.PasteToFillSlideId + TextCollection.MenuChart,
        TextCollection.PasteToFillSlideId + TextCollection.MenuTable,
        TextCollection.PasteToFillSlideId + TextCollection.MenuTableCell,
        TextCollection.PasteToFillSlideId + TextCollection.MenuSlide,
        TextCollection.PasteToFillSlideId + TextCollection.MenuSmartArt,
        TextCollection.PasteToFillSlideId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteToFillSlideId + TextCollection.MenuEditSmartArtText,
        TextCollection.PasteToFillSlideId + TextCollection.RibbonButton)]
    class PasteToFillSlideEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !Graphics.IsClipboardEmpty();
        }
    }
}
