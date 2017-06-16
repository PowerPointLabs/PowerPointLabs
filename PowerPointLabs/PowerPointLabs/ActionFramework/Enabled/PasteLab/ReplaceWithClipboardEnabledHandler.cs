using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Enabled.PasteLab
{
    [ExportEnabledRibbonId(
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuShape,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuLine,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuFreeform,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuPicture,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuGroup,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuInk,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuVideo,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuChart,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuTable,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuTableCell,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuSmartArt,
        TextCollection.ReplaceWithClipboardId + TextCollection.RibbonButton)]
    class ReplaceWithClipboardEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !Graphics.IsClipboardEmpty() && IsSelectionSingleShape();
        }
    }
}