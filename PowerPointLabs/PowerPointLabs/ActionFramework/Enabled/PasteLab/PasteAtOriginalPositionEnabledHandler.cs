using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Enabled.PasteLab
{
    [ExportEnabledRibbonId(
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuShape,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuLine,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuFreeform,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuPicture,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuGroup,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuInk,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuVideo,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuTextEdit,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuChart,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuTable,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuTableCell,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuSlide,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuSmartArt,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteAtOriginalPositionId + TextCollection.MenuEditSmartArtText,
        TextCollection.PasteAtOriginalPositionId + TextCollection.RibbonButton)]
    class PasteAtOriginalPositionEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !Graphics.IsClipboardEmpty();
        }
    }
}