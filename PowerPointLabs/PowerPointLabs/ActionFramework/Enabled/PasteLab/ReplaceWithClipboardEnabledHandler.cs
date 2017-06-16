using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Enabled.PasteLab
{
    [ExportEnabledRibbonId(
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuShape,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuLine,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuFreeform,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuPicture,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuGroup,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuInk,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuVideo,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuChart,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuTable,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuTableCell,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuSmartArt,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.RibbonButton)]
    class ReplaceWithClipboardEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return !Graphics.IsClipboardEmpty() && IsSelectionSingleShape();
        }
    }
}