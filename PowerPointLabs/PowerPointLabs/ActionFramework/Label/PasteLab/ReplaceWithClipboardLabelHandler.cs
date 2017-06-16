using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
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
    class ReplaceWithClipboardLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.ReplaceWithClipboard;
        }
    }
}